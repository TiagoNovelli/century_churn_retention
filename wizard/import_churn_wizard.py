import base64
import io
from datetime import date

from odoo import _, api, fields, models
from odoo.exceptions import UserError

try:
    import openpyxl
except ImportError:
    openpyxl = None


class ImportChurnWizard(models.TransientModel):
    _name = "century.import.churn.wizard"
    _description = "Importar Lista de Churn"

    file_data = fields.Binary(string="Arquivo Excel (.xlsx)", required=True)
    file_name = fields.Char(string="Nome do Arquivo")

    coordenador_id = fields.Many2one(
        "res.users",
        string="Coordenador Responsavel",
        required=True,
        default=lambda self: self.env.user,
    )
    team_id = fields.Many2one("crm.team", string="Equipe de Vendas")
    date_deadline = fields.Date(
        string="Prazo para Contato",
        default=lambda self: date.today().replace(day=28),
    )
    import_batch = fields.Char(
        string="Nome do Lote",
        default=lambda self: f'Churn {date.today().strftime("%Y-%m")}',
        required=True,
        help="Identificador do lote para rastrear esta importacao",
    )

    duplicates_action = fields.Selection(
        [
            ("skip", "Ignorar duplicatas"),
            ("update", "Atualizar probabilidade se ja existir"),
        ],
        string="Se cliente ja estiver no pipeline",
        default="update",
        required=True,
    )
    only_ab = fields.Boolean(
        string="Importar apenas Curva A e B",
        default=True,
        help="Recomendado: curva C tem churn estruturalmente alto e baixo ROI de CS",
    )

    preview_html = fields.Html(string="Preview", readonly=True)
    total_rows = fields.Integer(string="Total de linhas", readonly=True)
    valid_rows = fields.Integer(string="Linhas validas", readonly=True)
    invalid_rows = fields.Integer(string="Linhas com erro", readonly=True)

    EXPECTED_COLS = {
        "cliente": ["cliente", "nome", "razao_social", "name"],
        "cnpj": ["cnpj", "cpf", "vat", "cnpj_cpf"],
        "curva": ["curva", "curva_abc", "abc"],
        "prob_churn": ["prob_churn", "probabilidade", "churn_prob", "probabilidade_churn"],
        "risco": ["risco", "nivel_risco", "risk"],
        "receita_total": ["receita_total", "receita", "revenue"],
        "n_pedidos": ["n_pedidos", "pedidos", "orders"],
        "recencia": ["recencia", "recencia_meses"],
        "var_receita": ["var_receita", "variacao_receita"],
    }

    def _find_col(self, headers, candidates):
        headers_clean = [str(h).strip().lower().replace(" ", "_") for h in headers]
        for candidate in candidates:
            if candidate.lower() in headers_clean:
                return headers_clean.index(candidate.lower())
        return None

    def _clean_cnpj(self, value):
        if not value:
            return ""
        return "".join(filter(str.isdigit, str(value)))

    def _find_partner(self, cnpj_val="", nome_val=""):
        Partner = self.env["res.partner"]
        partner = Partner.browse()

        candidates = []
        if cnpj_val:
            candidates.append(cnpj_val)
            if len(cnpj_val) > 14:
                candidates.append(cnpj_val[-14:])
            candidates.append(cnpj_val.lstrip("0"))

        seen = set()
        normalized_candidates = []
        for candidate in candidates:
            if candidate and candidate not in seen:
                normalized_candidates.append(candidate)
                seen.add(candidate)

        for candidate in normalized_candidates:
            partner = Partner.search([("vat", "ilike", candidate)], limit=1)
            if partner:
                return partner.commercial_partner_id

        if normalized_candidates:
            shortlist = Partner.search([("vat", "!=", False)])
            for record in shortlist:
                vat_digits = self._clean_cnpj(record.vat)
                if vat_digits in normalized_candidates:
                    return record.commercial_partner_id
                if len(vat_digits) > 14 and vat_digits[-14:] in normalized_candidates:
                    return record.commercial_partner_id

        if nome_val:
            partner = Partner.search([("name", "ilike", nome_val)], limit=1)
            if partner:
                return partner.commercial_partner_id

        return Partner.browse()

    def _parse_file(self):
        if not openpyxl:
            raise UserError(_("Biblioteca openpyxl nao instalada."))

        if not self.file_data:
            raise UserError(_("Nenhum arquivo selecionado."))

        try:
            data = base64.b64decode(self.file_data)
            workbook = openpyxl.load_workbook(
                io.BytesIO(data), read_only=True, data_only=True
            )
            worksheet = workbook.active
        except Exception as exc:
            raise UserError(_("Erro ao ler o arquivo: %s") % str(exc)) from exc

        rows = list(worksheet.iter_rows(values_only=True))
        if len(rows) < 2:
            raise UserError(_("O arquivo esta vazio ou nao tem dados alem do cabecalho."))

        headers = rows[0]
        col_map = {}
        for field_name, candidates in self.EXPECTED_COLS.items():
            col_map[field_name] = self._find_col(headers, candidates)

        if col_map.get("cnpj") is None and col_map.get("cliente") is None:
            raise UserError(
                _(
                    "O arquivo precisa ter ao menos uma coluna de identificacao do cliente "
                    "(CNPJ ou Nome). Colunas encontradas: %s"
                )
                % ", ".join([str(h) for h in headers])
            )

        return rows[1:], col_map

    @api.onchange("file_data", "only_ab")
    def _onchange_file_preview(self):
        if not self.file_data:
            self.preview_html = False
            self.total_rows = 0
            self.valid_rows = 0
            self.invalid_rows = 0
            return

        try:
            rows, col_map = self._parse_file()
        except UserError as exc:
            self.preview_html = f'<p class="text-danger">{exc.args[0]}</p>'
            self.total_rows = 0
            self.valid_rows = 0
            self.invalid_rows = 0
            return

        total = len(rows)
        valid = 0
        invalid = 0
        preview_rows = []

        for row in rows[:10]:
            cnpj_val = self._clean_cnpj(
                row[col_map["cnpj"]] if col_map.get("cnpj") is not None else ""
            )
            nome_val = row[col_map["cliente"]] if col_map.get("cliente") is not None else ""
            curva = (
                str(row[col_map["curva"]]).strip().upper()
                if col_map.get("curva") is not None and row[col_map["curva"]] is not None
                else ""
            )
            prob = row[col_map["prob_churn"]] if col_map.get("prob_churn") is not None else 0

            partner = self._find_partner(cnpj_val, str(nome_val or "").strip())

            status = "OK" if partner else "Nao encontrado"
            if self.only_ab and curva not in ("A", "B"):
                status = "Ignorado (Curva C)"

            preview_rows.append(
                f"""
                <tr>
                    <td>{nome_val or ""}</td>
                    <td>{cnpj_val}</td>
                    <td>{curva}</td>
                    <td>{prob or 0}</td>
                    <td>{status}</td>
                </tr>
                """
            )

            if partner:
                valid += 1
            else:
                invalid += 1

        self.total_rows = total
        self.valid_rows = valid
        self.invalid_rows = invalid
        self.preview_html = f"""
            <p><b>{total}</b> linhas encontradas |
               <b class="text-success">{valid}</b> clientes identificados |
               <b class="text-warning">{invalid}</b> nao encontrados (primeiras 10)</p>
            <table class="table table-sm table-bordered">
                <thead class="table-light">
                    <tr>
                        <th>Cliente</th><th>CNPJ</th><th>Curva</th>
                        <th>Prob. Churn</th><th>Status</th>
                    </tr>
                </thead>
                <tbody>{''.join(preview_rows)}</tbody>
            </table>
        """

    def action_import(self):
        self.ensure_one()
        rows, col_map = self._parse_file()

        stage_inicial = self.env["century.retention.stage"].search(
            [("sequence", "=", 1)], limit=1
        )
        if not stage_inicial:
            raise UserError(_("Nenhum estagio inicial encontrado. Configure os estagios primeiro."))

        criados = 0
        atualizados = 0
        ignorados = 0
        erros = []

        def get_val(row, field_name, default=None):
            idx = col_map.get(field_name)
            if idx is None:
                return default
            value = row[idx]
            return value if value is not None else default

        for i, row in enumerate(rows, start=2):
            try:
                curva = str(get_val(row, "curva", "")).strip().upper()
                if self.only_ab and curva not in ("A", "B"):
                    ignorados += 1
                    continue

                cnpj_val = self._clean_cnpj(get_val(row, "cnpj", ""))
                nome_val = str(get_val(row, "cliente", "") or "").strip()

                partner = self._find_partner(cnpj_val, nome_val)

                if not partner:
                    erros.append(
                        f'Linha {i}: cliente "{nome_val}" / CNPJ "{cnpj_val}" nao encontrado no Odoo'
                    )
                    ignorados += 1
                    continue

                prob_raw = float(get_val(row, "prob_churn", 0) or 0)
                prob_churn = prob_raw * 100 if prob_raw <= 1.0 else prob_raw

                vals = {
                    "partner_id": partner.id,
                    "curva_abc": curva if curva in ("A", "B", "C") else False,
                    "prob_churn": prob_churn,
                    "receita_total": float(get_val(row, "receita_total", 0) or 0),
                    "n_pedidos": int(get_val(row, "n_pedidos", 0) or 0),
                    "recencia": int(get_val(row, "recencia", 0) or 0),
                    "var_receita": float(get_val(row, "var_receita", 0) or 0),
                    "coordenador_id": self.coordenador_id.id,
                    "team_id": self.team_id.id if self.team_id else False,
                    "date_deadline": self.date_deadline,
                    "import_batch": self.import_batch,
                    "stage_id": stage_inicial.id,
                    "resultado": "em_processo",
                }

                if partner.user_id:
                    vals["representante_id"] = partner.user_id.id

                existing = self.env["century.retention.lead"].search(
                    [("partner_id", "=", partner.id), ("resultado", "=", "em_processo")],
                    limit=1,
                )

                if existing:
                    if self.duplicates_action == "update":
                        existing.write(
                            {
                                "prob_churn": vals["prob_churn"],
                                "curva_abc": vals["curva_abc"],
                                "receita_total": vals["receita_total"],
                                "import_batch": vals["import_batch"],
                            }
                        )
                        atualizados += 1
                    else:
                        ignorados += 1
                else:
                    self.env["century.retention.lead"].create(vals)
                    criados += 1

            except Exception as exc:
                erros.append(f"Linha {i}: {str(exc)}")

        action = self.env["ir.actions.actions"]._for_xml_id(
            "century_churn_retention.action_retention_leads"
        )
        action["name"] = f"Importacao: {self.import_batch}"
        action["domain"] = [("import_batch", "=", self.import_batch)]
        action["context"] = {
            "search_default_em_processo": 1,
            "search_default_curva_a": 1,
            "default_import_batch": self.import_batch,
        }

        if erros:
            action["help"] = (
                f"<p><b>{criados}</b> criados, <b>{atualizados}</b> atualizados, "
                f"<b>{ignorados}</b> ignorados.</p>"
                f"<p>Primeiros erros: {' | '.join(erros[:5])}</p>"
            )

        return action

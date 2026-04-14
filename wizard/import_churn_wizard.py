import base64
import io
from datetime import date

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError

try:
    import openpyxl
except ImportError:
    openpyxl = None


class ImportChurnWizard(models.TransientModel):
    _name = 'century.import.churn.wizard'
    _description = 'Importar Lista de Churn'

    # ── Arquivo ──────────────────────────────────────────────────────────────
    file_data = fields.Binary(string='Arquivo Excel (.xlsx)', required=True)
    file_name = fields.Char(string='Nome do Arquivo')

    # ── Configurações da importação ──────────────────────────────────────────
    coordenador_id = fields.Many2one(
        'res.users',
        string='Coordenador Responsável',
        required=True,
        default=lambda self: self.env.user,
    )
    team_id = fields.Many2one(
        'crm.team',
        string='Equipe de Vendas',
    )
    date_deadline = fields.Date(
        string='Prazo para Contato',
        default=lambda self: date.today().replace(day=28),
    )
    import_batch = fields.Char(
        string='Nome do Lote',
        default=lambda self: f'Churn {date.today().strftime("%Y-%m")}',
        required=True,
        help='Identificador do lote para rastrear esta importação',
    )

    # ── Comportamento ─────────────────────────────────────────────────────────
    duplicates_action = fields.Selection([
        ('skip',   'Ignorar duplicatas'),
        ('update', 'Atualizar probabilidade se já existir'),
    ], string='Se cliente já estiver no pipeline', default='update', required=True)

    only_ab = fields.Boolean(
        string='Importar apenas Curva A e B',
        default=True,
        help='Recomendado: curva C tem churn estruturalmente alto e baixo ROI de CS',
    )

    # ── Preview ───────────────────────────────────────────────────────────────
    preview_html = fields.Html(string='Preview', readonly=True)
    total_rows    = fields.Integer(string='Total de linhas', readonly=True)
    valid_rows    = fields.Integer(string='Linhas válidas', readonly=True)
    invalid_rows  = fields.Integer(string='Linhas com erro', readonly=True)

    # ──────────────────────────────────────────────────────────────────────────
    # Colunas esperadas no Excel
    # ──────────────────────────────────────────────────────────────────────────
    EXPECTED_COLS = {
        'cliente':       ['cliente', 'nome', 'razao_social', 'name'],
        'cnpj':          ['cnpj', 'cpf', 'vat', 'cnpj_cpf'],
        'curva':         ['curva', 'curva_abc', 'abc'],
        'prob_churn':    ['prob_churn', 'probabilidade', 'churn_prob', 'probabilidade_churn'],
        'risco':         ['risco', 'nivel_risco', 'risk'],
        'receita_total': ['receita_total', 'receita', 'revenue'],
        'n_pedidos':     ['n_pedidos', 'pedidos', 'orders'],
        'recencia':      ['recencia', 'recencia_meses'],
        'var_receita':   ['var_receita', 'variacao_receita'],
    }

    # ──────────────────────────────────────────────────────────────────────────
    # Helpers
    # ──────────────────────────────────────────────────────────────────────────

    def _find_col(self, headers, candidates):
        """Encontra a coluna no header ignorando case e espaços."""
        headers_clean = [str(h).strip().lower().replace(' ', '_') for h in headers]
        for c in candidates:
            if c.lower() in headers_clean:
                return headers_clean.index(c.lower())
        return None

    def _clean_cnpj(self, val):
        if not val:
            return ''
        return ''.join(filter(str.isdigit, str(val)))

    def _parse_file(self):
        if not openpyxl:
            raise UserError(_('Biblioteca openpyxl não instalada. Execute: pip install openpyxl'))

        if not self.file_data:
            raise UserError(_('Nenhum arquivo selecionado.'))

        try:
            data = base64.b64decode(self.file_data)
            wb   = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            ws   = wb.active
        except Exception as e:
            raise UserError(_('Erro ao ler o arquivo: %s') % str(e))

        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            raise UserError(_('O arquivo está vazio ou não tem dados além do cabeçalho.'))

        headers = rows[0]
        col_map = {}
        for field, candidates in self.EXPECTED_COLS.items():
            idx = self._find_col(headers, candidates)
            col_map[field] = idx

        # CNPJ ou Nome são obrigatórios para vincular ao parceiro
        if col_map.get('cnpj') is None and col_map.get('cliente') is None:
            raise UserError(_(
                'O arquivo precisa ter ao menos uma coluna de identificação do cliente '
                '(CNPJ ou Nome). Colunas encontradas: %s'
            ) % ', '.join([str(h) for h in headers]))

        return rows[1:], col_map

    # ──────────────────────────────────────────────────────────────────────────
    # Preview
    # ──────────────────────────────────────────────────────────────────────────

    @api.onchange('file_data', 'only_ab')
    def _onchange_file_preview(self):
        if not self.file_data:
            self.preview_html = False
            return

        try:
            rows, col_map = self._parse_file()
        except UserError as e:
            self.preview_html = f'<p class="text-danger">{e.args[0]}</p>'
            return

        total   = len(rows)
        valid   = 0
        invalid = 0
        preview_rows = []

        for row in rows[:10]:  # preview primeiras 10 linhas
            cnpj_val = self._clean_cnpj(
                row[col_map['cnpj']] if col_map.get('cnpj') is not None else ''
            )
            nome_val = row[col_map['cliente']] if col_map.get('cliente') is not None else ''
            curva    = str(row[col_map['curva']]).strip().upper() if col_map.get('curva') is not None else ''
            prob     = row[col_map['prob_churn']] if col_map.get('prob_churn') is not None else 0

            # Verifica se existe parceiro
            partner = False
            if cnpj_val:
                partner = self.env['res.partner'].search([('vat', 'like', cnpj_val)], limit=1)
            if not partner and nome_val:
                partner = self.env['res.partner'].search([('name', 'ilike', str(nome_val))], limit=1)

            status = '✅' if partner else '⚠️ Não encontrado'
            if self.only_ab and curva not in ('A', 'B'):
                status = '⏭️ Ignorado (Curva C)'

            preview_rows.append(f'''
                <tr>
                    <td>{nome_val}</td>
                    <td>{cnpj_val}</td>
                    <td>{curva}</td>
                    <td>{prob}</td>
                    <td>{status}</td>
                </tr>
            ''')

            if partner:
                valid += 1
            else:
                invalid += 1

        self.total_rows   = total
        self.valid_rows   = valid
        self.invalid_rows = invalid

        self.preview_html = f'''
            <p><b>{total}</b> linhas encontradas | 
               <b class="text-success">{valid}</b> clientes identificados | 
               <b class="text-warning">{invalid}</b> não encontrados (primeiras 10)</p>
            <table class="table table-sm table-bordered">
                <thead class="table-light">
                    <tr>
                        <th>Cliente</th><th>CNPJ</th><th>Curva</th>
                        <th>Prob. Churn</th><th>Status</th>
                    </tr>
                </thead>
                <tbody>{''.join(preview_rows)}</tbody>
            </table>
        '''

    # ──────────────────────────────────────────────────────────────────────────
    # Importação
    # ──────────────────────────────────────────────────────────────────────────

    def action_import(self):
        self.ensure_one()
        rows, col_map = self._parse_file()

        stage_inicial = self.env['century.retention.stage'].search(
            [('sequence', '=', 1)], limit=1
        )
        if not stage_inicial:
            raise UserError(_('Nenhum estágio inicial encontrado. Configure os estágios primeiro.'))

        criados    = 0
        atualizados = 0
        ignorados  = 0
        erros      = []

        def get_val(row, field, default=None):
            idx = col_map.get(field)
            if idx is None:
                return default
            v = row[idx]
            return v if v is not None else default

        for i, row in enumerate(rows, start=2):
            try:
                curva = str(get_val(row, 'curva', '')).strip().upper()

                # Filtra curva C se configurado
                if self.only_ab and curva not in ('A', 'B'):
                    ignorados += 1
                    continue

                # Localiza parceiro
                cnpj_val = self._clean_cnpj(get_val(row, 'cnpj', ''))
                nome_val = str(get_val(row, 'cliente', '') or '').strip()

                partner = False
                if cnpj_val:
                    partner = self.env['res.partner'].search(
                        [('vat', 'like', cnpj_val), ('is_company', '=', True)], limit=1
                    )
                if not partner and nome_val:
                    partner = self.env['res.partner'].search(
                        [('name', 'ilike', nome_val), ('is_company', '=', True)], limit=1
                    )

                if not partner:
                    erros.append(f'Linha {i}: cliente "{nome_val}" / CNPJ "{cnpj_val}" não encontrado no Odoo')
                    ignorados += 1
                    continue

                # Probabilidade — converte para % se estiver em decimal (0-1)
                prob_raw = float(get_val(row, 'prob_churn', 0) or 0)
                prob_churn = prob_raw * 100 if prob_raw <= 1.0 else prob_raw

                # Monta valores
                vals = {
                    'partner_id':    partner.id,
                    'curva_abc':     curva if curva in ('A', 'B', 'C') else False,
                    'prob_churn':    prob_churn,
                    'receita_total': float(get_val(row, 'receita_total', 0) or 0),
                    'n_pedidos':     int(get_val(row, 'n_pedidos', 0) or 0),
                    'recencia':      int(get_val(row, 'recencia', 0) or 0),
                    'var_receita':   float(get_val(row, 'var_receita', 0) or 0),
                    'coordenador_id': self.coordenador_id.id,
                    'team_id':       self.team_id.id if self.team_id else False,
                    'date_deadline': self.date_deadline,
                    'import_batch':  self.import_batch,
                    'stage_id':      stage_inicial.id,
                    'resultado':     'em_processo',
                }

                # Vincula representante se o parceiro tiver user responsável
                if partner.user_id:
                    vals['representante_id'] = partner.user_id.id

                # Verifica duplicatas
                existing = self.env['century.retention.lead'].search([
                    ('partner_id', '=', partner.id),
                    ('resultado', '=', 'em_processo'),
                ], limit=1)

                if existing:
                    if self.duplicates_action == 'update':
                        existing.write({
                            'prob_churn':    vals['prob_churn'],
                            'curva_abc':     vals['curva_abc'],
                            'receita_total': vals['receita_total'],
                            'import_batch':  vals['import_batch'],
                        })
                        atualizados += 1
                    else:
                        ignorados += 1
                else:
                    self.env['century.retention.lead'].create(vals)
                    criados += 1

            except Exception as e:
                erros.append(f'Linha {i}: {str(e)}')

        # ── Mensagem de resultado ──────────────────────────────────────────
        msg_parts = [
            f'✅ {criados} clientes criados',
            f'🔄 {atualizados} atualizados',
            f'⏭️ {ignorados} ignorados',
        ]
        if erros:
            msg_parts.append(f'⚠️ {len(erros)} erros: ' + ' | '.join(erros[:5]))

        # Notificação simples no chatter do wizard — compatível com Odoo 18
        tipo = 'success' if not erros else 'warning'

        # Redireciona para o pipeline filtrado pelo lote
        return {
            'type': 'ir.actions.act_window',
            'name': f'Importação: {self.import_batch}',
            'res_model': 'century.retention.lead',
            'view_mode': 'kanban,list,form',
            'views': [(False, 'kanban'), (False, 'list'), (False, 'form')],
            'domain': [('import_batch', '=', self.import_batch)],
            'context': {
                'search_default_em_processo': 1,
                'search_default_curva_a': 1,
                'import_result_message': ' | '.join(msg_parts),
            },
            'target': 'current',
        }

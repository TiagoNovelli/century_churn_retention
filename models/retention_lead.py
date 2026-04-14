from odoo import models, fields, api, _
from odoo.exceptions import UserError


class RetentionLead(models.Model):
    _name = 'century.retention.lead'
    _description = 'Cliente em Risco de Churn'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'prob_churn desc, curva_abc asc'
    _rec_name = 'partner_id'

    # ── Parceiro ────────────────────────────────────────────────────────────
    partner_id = fields.Many2one(
        'res.partner',
        string='Cliente',
        required=True,
        tracking=True,
        domain=[('is_company', '=', True)],
    )
    cnpj = fields.Char(
        related='partner_id.vat',
        string='CNPJ',
        store=True,
        readonly=True,
    )

    # ── Dados do modelo de churn ────────────────────────────────────────────
    prob_churn = fields.Float(
        string='Probabilidade de Churn (%)',
        digits=(5, 2),
        tracking=True,
        help='Probabilidade gerada pelo modelo preditivo (0-100)',
    )
    curva_abc = fields.Selection([
        ('A', 'Curva A — Alto Valor'),
        ('B', 'Curva B — Médio Valor'),
        ('C', 'Curva C — Baixo Valor'),
    ], string='Curva ABC', tracking=True)

    nivel_risco = fields.Selection([
        ('alto',  '🔴 Alto'),
        ('medio', '🟡 Médio'),
        ('baixo', '🟢 Baixo'),
    ], string='Nível de Risco', tracking=True, compute='_compute_nivel_risco', store=True)

    receita_total = fields.Monetary(
        string='Receita Último Ano',
        currency_field='currency_id',
    )
    n_pedidos = fields.Integer(string='Pedidos no Ano')
    recencia  = fields.Integer(string='Recência (meses)')
    var_receita = fields.Float(string='Variação Receita (%)', digits=(5, 2))

    # ── Responsáveis ────────────────────────────────────────────────────────
    representante_id = fields.Many2one(
        'res.users',
        string='Vendedor',
        related='partner_id.user_id',
        store=True,
        readonly=True,
    )
    team_id = fields.Many2one(
        'crm.team',
        string='Equipe de Vendas',
        compute='_compute_sales_team',
        store=True,
        readonly=True,
    )
    team_leader_id = fields.Many2one(
        'res.users',
        string='Lider de Vendas',
        related='team_id.user_id',
        store=True,
        readonly=True,
    )

    # ── Pipeline ────────────────────────────────────────────────────────────
    stage_id = fields.Many2one(
        'century.retention.stage',
        string='Estágio',
        required=True,
        tracking=True,
        default=lambda self: self._default_stage(),
        group_expand='_read_group_stage_ids',
    )
    kanban_state = fields.Selection([
        ('normal',   'Em Progresso'),
        ('done',     'Pronto para Avançar'),
        ('blocked',  'Bloqueado'),
    ], string='Estado Kanban', default='normal', tracking=True)

    # ── Datas ────────────────────────────────────────────────────────────────
    date_import   = fields.Date(string='Data de Importação', default=fields.Date.today)
    date_contact  = fields.Date(string='Data do Contato', tracking=True)
    date_deadline = fields.Date(string='Prazo', tracking=True)

    # ── Resultado ────────────────────────────────────────────────────────────
    resultado = fields.Selection([
        ('recuperado',  '✅ Recuperado'),
        ('churned',     '❌ Churned'),
        ('em_processo', '⏳ Em Processo'),
    ], string='Resultado', default='em_processo', tracking=True)

    motivo_churn = fields.Selection([
        ('sell_out',      'Sell-out fraco'),
        ('concorrencia',  'Concorrência'),
        ('preco',         'Preço'),
        ('atendimento',   'Atendimento'),
        ('mix',           'Mix inadequado'),
        ('fechamento',    'Fechamento da loja'),
        ('outro',         'Outro'),
    ], string='Motivo do Churn', tracking=True)

    observacoes = fields.Text(string='Observações')

    # ── Financeiro ────────────────────────────────────────────────────────────
    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id,
    )

    # ── Prioridade ────────────────────────────────────────────────────────────
    priority = fields.Selection([
        ('0', 'Normal'),
        ('1', '⭐ Importante'),
        ('2', '⭐⭐ Urgente'),
        ('3', '⭐⭐⭐ Crítico'),
    ], string='Prioridade', default='0')

    # ── Lote de importação ────────────────────────────────────────────────────
    import_batch = fields.Char(string='Lote de Importação')

    # ──────────────────────────────────────────────────────────────────────────
    # Computed
    # ──────────────────────────────────────────────────────────────────────────

    @api.depends('prob_churn')
    def _compute_nivel_risco(self):
        for rec in self:
            if rec.prob_churn >= 70:
                rec.nivel_risco = 'alto'
            elif rec.prob_churn >= 40:
                rec.nivel_risco = 'medio'
            else:
                rec.nivel_risco = 'baixo'

    @api.depends('representante_id')
    def _compute_sales_team(self):
        TeamMember = self.env['crm.team.member']
        for rec in self:
            rec.team_id = False
            if not rec.representante_id:
                continue
            member = TeamMember.search(
                [('user_id', '=', rec.representante_id.id)],
                order='id asc',
                limit=1,
            )
            if member:
                rec.team_id = member.crm_team_id

    def _default_stage(self):
        return self.env['century.retention.stage'].search(
            [('sequence', '=', 1)], limit=1
        )

    @api.model
    def _read_group_stage_ids(self, stages, domain, order=None):
        return self.env['century.retention.stage'].search([], order='sequence asc')

    # ──────────────────────────────────────────────────────────────────────────
    # Regras de visibilidade
    # ──────────────────────────────────────────────────────────────────────────

    @api.model
    def _is_coordenador(self):
        return self.env.user.has_group('century_churn_retention.group_retention_coordinator')

    # ──────────────────────────────────────────────────────────────────────────
    # Actions
    # ──────────────────────────────────────────────────────────────────────────

    def action_marcar_contato(self):
        self.write({
            'date_contact': fields.Date.today(),
            'kanban_state': 'normal',
        })
        self.message_post(body=_('Contato realizado em %s') % fields.Date.today())

    def action_recuperado(self):
        stage_recuperado = self.env['century.retention.stage'].search(
            [('is_won', '=', True)], limit=1
        )
        self.write({
            'resultado': 'recuperado',
            'stage_id': stage_recuperado.id if stage_recuperado else self.stage_id.id,
        })
        self.message_post(body=_('✅ Cliente marcado como RECUPERADO'))

    def action_churned(self):
        stage_lost = self.env['century.retention.stage'].search(
            [('is_lost', '=', True)], limit=1
        )
        self.write({
            'resultado': 'churned',
            'stage_id': stage_lost.id if stage_lost else self.stage_id.id,
        })
        self.message_post(body=_('❌ Cliente marcado como CHURNED'))


class RetentionStage(models.Model):
    _name = 'century.retention.stage'
    _description = 'Estágio do Pipeline de Retenção'
    _order = 'sequence asc'

    name     = fields.Char(string='Nome', required=True)
    sequence = fields.Integer(string='Sequência', default=10)
    is_won   = fields.Boolean(string='Estágio de Recuperação')
    is_lost  = fields.Boolean(string='Estágio de Churn Confirmado')
    fold     = fields.Boolean(string='Dobrado no Kanban')
    color    = fields.Integer(string='Cor')
    description = fields.Text(string='Descrição / Critério de entrada')

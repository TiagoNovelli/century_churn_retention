{
    'name': 'Century — Retenção de Clientes (Churn)',
    'version': '18.0.1.0.0',
    'category': 'Sales/CRM',
    'summary': 'Pipeline de retenção B2B com upload de lista de churn',
    'description': """
        Módulo de gestão de retenção de clientes Century Sofás.
        - Upload de lista de clientes em risco via Excel
        - Pipeline de retenção com estágios customizados
        - Visibilidade por representante e coordenador
        - Integração com res.partner via CNPJ
    """,
    'author': 'Century Sofás',
    'depends': ['crm', 'sale_management', 'mail'],
    'data': [
        'security/security.xml',
        'security/ir.model.access.csv',
        'data/crm_stage_data.xml',
        'views/retention_views.xml',
        'views/retention_menu.xml',
        'wizard/import_churn_views.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}

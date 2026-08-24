# -*- coding: utf-8 -*-
from datetime import datetime
from zoneinfo import ZoneInfo
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

FUSO_BRASIL = ZoneInfo("America/Sao_Paulo")


def hoje_brasil():
    """Data atual no fuso de Brasilia. O servidor roda em UTC, entao usar
    date.today() diretamente pode adiantar o dia em ate 3 horas."""
    return datetime.now(FUSO_BRASIL).date()

CATEGORIAS_ENTRADA = [
    "Venda Dinheiro",
    "Venda Cartao Debito",
    "Venda Cartao Credito",
    "Venda PIX",
    "Suprimento de Caixa",
    "Outras Entradas",
]
CATEGORIAS_SAIDA = [
    "Sangria",
    "Pagamento Fornecedor",
    "Despesa Operacional",
    "Outras Saidas",
]
CATEGORIAS_MOVIMENTO = CATEGORIAS_ENTRADA + CATEGORIAS_SAIDA

CATEGORIAS_PAGAR = [
    "Mercadorias/Fornecedores", "Aluguel", "Salarios e Encargos", "Energia Eletrica",
    "Agua", "Telefone/Internet", "Impostos", "Manutencao", "Embalagens", "Outros",
]
CATEGORIAS_RECEBER = [
    "Cartao de Credito", "Cartao de Debito", "Convenio/Fiado", "Cheque Pre-datado", "Boleto", "Outros",
]

MESES_PT = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


class Config(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    saldo_inicial = db.Column(db.Float, default=0.0)
    data_inicial = db.Column(db.Date, default=hoje_brasil)


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    ativo = db.Column(db.Boolean, default=True)
    admin = db.Column(db.Boolean, default=False)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)


class MovimentoCaixa(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.Date, nullable=False, default=hoje_brasil)
    descricao = db.Column(db.String(200), nullable=False)
    categoria = db.Column(db.String(50), nullable=False)
    valor = db.Column(db.Float, nullable=False)

    @property
    def tipo(self):
        return "Entrada" if self.categoria in CATEGORIAS_ENTRADA else "Saida"

    @property
    def entrada(self):
        return self.valor if self.tipo == "Entrada" else 0.0

    @property
    def saida(self):
        return self.valor if self.tipo == "Saida" else 0.0


class ContaPagar(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fornecedor = db.Column(db.String(150), nullable=False)
    descricao = db.Column(db.String(200))
    categoria = db.Column(db.String(50), nullable=False)
    data_emissao = db.Column(db.Date)
    data_vencimento = db.Column(db.Date, nullable=False)
    valor = db.Column(db.Float, nullable=False)
    data_pagamento = db.Column(db.Date, nullable=True)
    observacoes = db.Column(db.String(300))
    movimento_id = db.Column(db.Integer, db.ForeignKey("movimento_caixa.id"), nullable=True)

    @property
    def status(self):
        if self.data_pagamento:
            return "Pago"
        if self.data_vencimento < hoje_brasil():
            return "Atrasado"
        return "Em Aberto"

    @property
    def dias_atraso(self):
        if self.data_pagamento or self.data_vencimento >= hoje_brasil():
            return 0
        return (hoje_brasil() - self.data_vencimento).days


class ContaReceber(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cliente = db.Column(db.String(150), nullable=False)
    descricao = db.Column(db.String(200))
    categoria = db.Column(db.String(50), nullable=False)
    data_venda = db.Column(db.Date)
    data_prevista = db.Column(db.Date, nullable=False)
    valor_bruto = db.Column(db.Float, nullable=False)
    taxa_percentual = db.Column(db.Float, default=0.0)
    data_recebimento = db.Column(db.Date, nullable=True)
    observacoes = db.Column(db.String(300))
    movimento_id = db.Column(db.Integer, db.ForeignKey("movimento_caixa.id"), nullable=True)

    @property
    def valor_liquido(self):
        return round(self.valor_bruto * (1 - (self.taxa_percentual or 0)), 2)

    @property
    def status(self):
        if self.data_recebimento:
            return "Recebido"
        if self.data_prevista < hoje_brasil():
            return "Atrasado"
        return "Em Aberto"

    @property
    def dias_atraso(self):
        if self.data_recebimento or self.data_prevista >= hoje_brasil():
            return 0
        return (hoje_brasil() - self.data_prevista).days


class DreMensal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ano = db.Column(db.Integer, nullable=False)
    mes = db.Column(db.Integer, nullable=False)  # 1-12
    receita_bruta = db.Column(db.Float, default=0.0)
    devolucoes = db.Column(db.Float, default=0.0)
    aliquota_impostos = db.Column(db.Float, default=0.04)
    cmv = db.Column(db.Float, default=0.0)
    salarios = db.Column(db.Float, default=0.0)
    aluguel = db.Column(db.Float, default=0.0)
    energia_agua = db.Column(db.Float, default=0.0)
    materiais_embalagens = db.Column(db.Float, default=0.0)
    marketing = db.Column(db.Float, default=0.0)
    outras_despesas = db.Column(db.Float, default=0.0)
    despesas_financeiras = db.Column(db.Float, default=0.0)
    receitas_financeiras = db.Column(db.Float, default=0.0)
    numero_vendas = db.Column(db.Integer, default=0)

    __table_args__ = (db.UniqueConstraint("ano", "mes", name="uq_dre_ano_mes"),)

    @property
    def impostos(self):
        return round(self.receita_bruta * self.aliquota_impostos, 2)

    @property
    def receita_liquida(self):
        return round(self.receita_bruta - self.devolucoes - self.impostos, 2)

    @property
    def lucro_bruto(self):
        return round(self.receita_liquida - self.cmv, 2)

    @property
    def margem_bruta(self):
        return (self.lucro_bruto / self.receita_liquida) if self.receita_liquida else 0.0

    @property
    def total_despesas_operacionais(self):
        return round(sum([self.salarios, self.aluguel, self.energia_agua,
                           self.materiais_embalagens, self.marketing, self.outras_despesas]), 2)

    @property
    def resultado_operacional(self):
        return round(self.lucro_bruto - self.total_despesas_operacionais, 2)

    @property
    def margem_operacional(self):
        return (self.resultado_operacional / self.receita_liquida) if self.receita_liquida else 0.0

    @property
    def resultado_liquido(self):
        return round(self.resultado_operacional - self.despesas_financeiras + self.receitas_financeiras, 2)

    @property
    def margem_liquida(self):
        return (self.resultado_liquido / self.receita_liquida) if self.receita_liquida else 0.0

    @property
    def ponto_equilibrio(self):
        return (self.total_despesas_operacionais / self.margem_bruta) if self.margem_bruta else 0.0

    @property
    def ticket_medio(self):
        return (self.receita_bruta / self.numero_vendas) if self.numero_vendas else 0.0

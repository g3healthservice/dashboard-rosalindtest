#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modelos editáveis (.docx) das peças de instrução da ação RosalindTest — NEUTROS (sem a G3),
padrão edital: Times New Roman 12, justificado, sem cores, margens ABNT (3/2/3/2 cm), páginas numeradas.
Rotas: MROSC (Lei 13.019/2014) e CPSI (LC 182/2021). Conteúdo alinhado ao gerador da plataforma.
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent / "modelos"
BASE.mkdir(exist_ok=True)
DISC = ("Documento em modelo de trabalho (minuta). Os campos entre colchetes devem ser preenchidos pelo órgão com os "
        "dados locais. As quantidades devem ser confirmadas pelo SISCAN, pelo e-SUS APS e pela base populacional; os "
        "valores, por pesquisa de mercado própria (IN SEGES/ME nº 65/2021). Não substitui o juízo da área técnica, da "
        "assessoria jurídica e do controle interno.")


def novo():
    doc = Document()
    f = doc.styles["Normal"].font; f.name = "Times New Roman"; f.size = Pt(12)
    doc.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    for s in doc.sections:
        s.top_margin, s.left_margin, s.bottom_margin, s.right_margin = Cm(3), Cm(3), Cm(2), Cm(2)
        p = s.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run("Página "); r.font.size = Pt(10)
        for code in ("begin", None, "end"):   # campo PAGE
            if code is None:
                it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"; r2 = p.add_run(); r2._r.append(it)
            else:
                fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), code); r2 = p.add_run(); r2._r.append(fc)
    return doc


def cab(doc, *linhas):
    for i, t in enumerate(("[ BRASÃO / TIMBRE DO ÓRGÃO ]", "[ SECRETARIA DE SAÚDE DE ______ ]") + linhas):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t); r.bold = i > 0
    p = doc.add_paragraph(); r = p.add_run(DISC); r.italic = True; r.font.size = Pt(10)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def h(doc, t):
    p = doc.add_paragraph(); r = p.add_run(t.upper()); r.bold = True


def par(doc, t):
    doc.add_paragraph(t).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def al(doc, itens):
    for i, t in enumerate(itens):
        p = doc.add_paragraph(f"{chr(97 + i)}) {t}"); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(1)


def tab(doc, head, rows):
    tb = doc.add_table(rows=1, cols=len(head)); tb.style = "Table Grid"
    for c, v in zip(tb.rows[0].cells, head):
        c.text = ""; c.paragraphs[0].add_run(v).bold = True
    for row in rows:
        for c, v in zip(tb.add_row().cells, row): c.text = v
    doc.add_paragraph()


def ass(doc, *blocos):
    for b in blocos:
        doc.add_paragraph()
        for t in ("_______________________________________",) + b:
            doc.add_paragraph(t).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("[ Local ], ____ de ______________ de 20__.").alignment = WD_ALIGN_PARAGRAPH.CENTER


def salvar(doc, nome):
    doc.save(BASE / nome); print("ok", nome)


LEG = [
    ("CF, art. 196; Lei nº 8.080/1990", "Saúde como direito; detecção precoce", "Fundamento da ação"),
    ("Lei nº 11.664/2008", "Mamografia a partir dos 40 anos no SUS", "A ação soma acesso; não substitui a mamografia"),
    ("Lei nº 14.758/2023 (PNPCC)", "Rastreamento com base em evidências; navegação", "Navegação até o diagnóstico; avaliação de resultados"),
    ("Lei nº 12.401/2011; Decreto nº 7.646/2011", "Incorporação via CONITEC, com registro Anvisa", "Teste in house não entra na Tabela SIGTAP enquanto LDT"),
    ("Lei nº 8.080/1990, art. 19-T", "Veda pagar procedimento experimental ou não autorizado", "Custeio fora da tabela SUS: contrapartida ou CPSI"),
    ("RDC Anvisa nº 978/2025 e alterações", "Requisitos dos serviços de análises clínicas", "Responsabilidade técnica do laboratório executor"),
    ("Lei nº 6.360/1976; RDC Anvisa nº 830/2023", "Registro de produtos para diagnóstico in vitro", "Exigível se virar kit comercial"),
    ("Lei nº 13.019/2014", "Parcerias com OSC; rede (art. 35-A)", "Rota A — MROSC"),
    ("LC nº 182/2021, arts. 12 a 15", "Licitação especial e CPSI (até 12+12 meses; até R$ 1,6 mi)", "Rota B — piloto de inovação"),
    ("Lei nº 14.133/2021, art. 74, I", "Inexigibilidade por exclusividade", "Não recomendada (não afasta o art. 19-T)"),
    ("Lei nº 12.732/2012; Lei nº 13.896/2019", "60 dias para tratar; 30 dias para exames na suspeita", "Prazo do fluxo de confirmação"),
    ("Lei nº 13.709/2018; Lei nº 14.874/2024", "Dado sensível; pesquisa com seres humanos", "TCLE e CEP/CONEP"),
]


def nota():
    d = novo(); cab(d, "NOTA TÉCNICA DE ENQUADRAMENTO LEGAL E DE CUSTEIO", "Ação de rastreamento do câncer de mama por biópsia líquida (RosalindTest, método in house)")
    h(d, "1. Objeto"); par(d, "Orientar o enquadramento jurídico, sanitário e orçamentário de ação de rastreamento do câncer de mama, complementar à mamografia, que utiliza o exame de sangue RosalindTest — metodologia in house, sem avaliação da CONITEC e sem procedimento na Tabela SIGTAP —, em [ município/UF ].")
    h(d, "2. Fundamentação técnica")
    al(d, ["Câncer de mama: 73.610 casos novos/ano e mais de 20 mil óbitos (INCA); ~40% diagnosticados em estágio avançado.",
           "População-alvo de [ ___ ] mulheres de 40 a 74 anos (IBGE, Censo 2022); cobertura estimada de [ __ ]% ante meta de 70% (INCA).",
           "Diretriz do Ministério da Saúde (set/2025): rastreamento bienal de 50 a 74 anos; acesso sob demanda de 40 a 49 anos.",
           "RosalindTest: PCR digital em duplex do mRNA de HIF-1alfa e GLUT1 (corte 5 cópias/µL); coorte de 750 mulheres; sensibilidade 90,65%, especificidade 96,47%; laudo em 15 a 20 dias; estudos iniciais.",
           "Com prevalência de ~5/1.000, o valor preditivo positivo é de ~11%: a maioria dos alterados não será câncer — confirmação por imagem obrigatória."])
    h(d, "3. O que diz a legislação vigente"); tab(d, ["Norma", "Conteúdo", "Consequência"], LEG)
    par(d, "Pesquisa normativa de novembro de 2026; conferir a redação vigente antes da assinatura.")
    h(d, "4. Natureza in house")
    al(d, ["Não há produto a adquirir: contrata-se serviço laboratorial dentro de ação de saúde.",
           "Responsabilidade sanitária do laboratório executor (RDC Anvisa nº 978/2025).",
           "LiqSci Biotecnologia (CNPJ 50.433.201/0001-50), titular da tecnologia, como interveniente anuente.",
           "Paradoxo do in house: o que permite executar já impede a incorporação (que exige registro de produto)."])
    h(d, "5. Rotas jurídicas")
    tab(d, ["Rota", "Instrumentos", "Atenção"], [
        ("A — MROSC (recomendada)", "Termo de Colaboração; Plano de Trabalho; Termo de Atuação em Rede; contrato laboratorial", "Eixo como contrapartida; chamamento ou dispensa motivada"),
        ("B — CPSI", "Edital de licitação especial; TR; CPSI", "Descrever o problema; teto R$ 1,6 mi; CEP"),
        ("C — Inexigibilidade", "Lei 14.133, art. 74, I", "Não recomendada")])
    h(d, "6. Fontes de custeio")
    tab(d, ["Fonte", "Admissível?"], [("Contrapartida de OSC/parceiro privado", "Sim — recomendada"), ("Recurso próprio", "Sim, fora da tabela"),
        ("Transferência especial (CF, art. 166-A)", "Sim"), ("Inovação (CPSI/FINEP/FAP)", "Sim"), ("MAC/PAP vinculados à tabela", "Não para o exame; sim para a confirmação")])
    h(d, "7. Salvaguardas")
    al(d, ["Complementar à mamografia; imagem em até 30 dias para alteradas, pactuada antes do início.", "TCLE informando que o exame não diagnostica.",
           "LGPD; indicadores agregados à gestão.", "CEP/CONEP para uso de dados em validação.", "Valor global da ação; nunca preço avulso por exame.",
           "Quem fornece a solução não executa com o mesmo recurso."])
    h(d, "8. Conclusão")
    par(d, "A ação é executável de forma lícita desde que: o exame seja serviço laboratorial in house dentro de ação de rastreamento; o custeio ocorra fora da tabela SUS (contrapartida ou CPSI), em observância ao art. 19-T da Lei nº 8.080/1990; a confirmação por imagem esteja pactuada; e a geração de evidência observe o CEP/CONEP.")
    ass(d, ("Área Técnica", "[ nome / cargo ]")); salvar(d, "NotaTecnica_RosalindTest_Enquadramento.docx")


def etp():
    d = novo(); cab(d, "ESTUDO TÉCNICO PRELIMINAR — ETP", "Ação de rastreamento do câncer de mama por biópsia líquida, complementar à mamografia")
    tab(d, ["Campo", "Conteúdo"], [("Processo nº", "[ ___ ]"), ("Unidade demandante", "[ ___ ]"), ("Rota", "[ MROSC — Lei 13.019/2014 | CPSI — LC 182/2021 ]"),
        ("Fonte", "[ contrapartida / recurso próprio / transferência especial / inovação ] — fora da tabela SUS")])
    for t, txt in [
        ("1. Necessidade", "Câncer de mama com alta mortalidade evitável; cobertura de [ __ ]% para [ ___ ] mulheres de 40 a 74 anos; barreiras de acesso à mamografia (deslocamento, filas, mamas densas, BI-RADS 0)."),
        ("2. Alinhamento", "Plano Municipal de Saúde/PAS [ previsão ]; Lei 11.664/2008; Lei 14.758/2023. Exame sem SIGTAP: custeio fora da tabela (art. 19-T da Lei 8.080/1990)."),
    ]: h(d, t); par(d, txt)
    h(d, "3. Requisitos")
    al(d, ["Laboratório com alvará, CNES, RT e POP do método (RDC Anvisa 978/2025); validação analítica documentada;", "Anuência da LiqSci (interveniente);",
           "Coleta por profissional habilitado, transporte refrigerado, PGRSS (RDC 222/2018);", "Laudo em até 20 dias; LGPD;",
           "Imagem para alteradas em até 30 dias (Lei 13.896/2019); navegação (Lei 14.758/2023); TCLE."])
    h(d, "4. Quantidades"); tab(d, ["Parâmetro", "Estimativa"], [("População-alvo 40–74", "[ ___ ]"), ("Convocadas", "[ ___ ]"), ("Exames", "[ ___ ]"), ("Alteradas → imagem", "[ ___ ]")])
    h(d, "5. Soluções"); par(d, "Mamografia (necessária para confirmação) e biópsia líquida RosalindTest (amplia acesso e prioriza a imagem; estudos iniciais).")
    h(d, "6. Valor"); tab(d, ["Componente", "Qtd.", "Unitário", "Total"], [("Serviço laboratorial in house", "[ ]", "[ R$ ]", "[ R$ ]"), ("Logística", "[ ]", "[ R$ ]", "[ R$ ]"), ("Valor global", "—", "—", "[ R$ ]")])
    for t, txt in [("7. Não-parcelamento", "Cadeia única coleta–processamento–laudo–navegação."),
                   ("8. Correlatas", "Retaguarda de imagem (SIGTAP/OCI) pactuada antes do início."),
                   ("9. Resultados", "Cobertura [ __ ]% → [ __ ]%; 100% das alteradas com imagem em 30 dias; relatório de desempenho (VPP, estadiamento)."),
                   ("10. Providências", "Fonte formalizada; fluxo de imagem; laboratório habilitado e anuência da LiqSci; TCLE e CEP."),
                   ("11. Impactos ambientais", "Resíduos biológicos e perfurocortantes conforme PGRSS."),
                   ("12. Conclusão", "Viável como ação complementar custeada fora da tabela SUS, pela rota indicada.")]:
        h(d, t); par(d, txt)
    ass(d, ("Responsável técnico", "[ nome / matrícula ]"), ("Autoridade competente", "")); salvar(d, "ETP_RosalindTest_RastreioMama.docx")


def plano():
    d = novo(); cab(d, "PLANO DE TRABALHO (LEI Nº 13.019/2014, ART. 22)", "Ação de rastreamento do câncer de mama — RosalindTest")
    tab(d, ["Papel", "Entidade"], [("Administração", "[ Secretaria ]"), ("OSC celebrante", "[ razão social / CNPJ ]"), ("OSC executante (rede)", "[ razão social / CNPJ — CNAE de saúde ]"),
        ("Fornecedores (anexo técnico)", "[ laboratório executor ] + LiqSci (interveniente)")])
    h(d, "1. Objeto"); par(d, "Ação de rastreamento do câncer de mama, complementar à mamografia: convocação, coleta de sangue, processamento in house, laudo, devolutiva e navegação até a imagem.")
    h(d, "2. Realidade e nexo (art. 22, I)"); par(d, "[ cobertura atual, população-alvo, barreiras locais ]")
    h(d, "3. Metas e etapas (art. 22, II)")
    tab(d, ["Etapa", "Atividade", "Prazo"], [("1", "Pactuação do fluxo de imagem e metas", "Mês 1"), ("2", "Capacitação", "Mês 1–2"), ("3", "Convocação", "Mês 2+"),
        ("4", "Coleta", "Contínuo"), ("5", "Análise in house e laudo (até 20 dias)", "Contínuo"), ("6", "Navegação à imagem (30 dias)", "Contínuo"), ("7", "Relatório de desempenho", "Fim do ciclo")])
    h(d, "4. Indicadores (art. 22, III)")
    al(d, ["Exames sobre a meta;", "Adesão;", "Alteradas com imagem em 30 dias (meta 100%);", "VPP e cânceres confirmados;", "Estadiamento ao diagnóstico;", "Tempo coleta–laudo."])
    h(d, "5. Receitas e despesas"); tab(d, ["Componente", "Memória", "Valor"], [("Serviço laboratorial", "[ qtd × R$ ]", "[ R$ ]"), ("Logística", "[ qtd × R$ ]", "[ R$ ]"), ("Valor global", "—", "[ R$ ]")])
    h(d, "6. Contrapartida"); par(d, "[ Opção A — eixo integral como contrapartida em bens/serviços (art. 35, § 1º) | Opção B — componente custeado, com justificativa técnica e anuência expressa ].")
    h(d, "7. Desembolso e prestação de contas"); par(d, "[ parcelas vinculadas a laudos e relatórios ]; conta específica (art. 51); prestação de contas (arts. 63 a 72).")
    ass(d, ("OSC celebrante", "Representante legal"), ("Administração Pública", "Gestor da parceria")); salvar(d, "MROSC_PlanoDeTrabalho_RosalindTest.docx")


def clausulas(d, itens):
    for n, (t, txt) in enumerate(itens, 1):
        h(d, f"Cláusula {n}ª — {t}"); par(d, txt)


def termo():
    d = novo(); cab(d, "TERMO DE COLABORAÇÃO Nº [ ___ ]", "Lei nº 13.019/2014 — ação de rastreamento do câncer de mama (RosalindTest)")
    par(d, "Celebram a [ Secretaria — CNPJ — representante ] (ADMINISTRAÇÃO) e [ OSC — CNPJ — representante ] (OSC), nos termos da Lei nº 13.019/2014:")
    clausulas(d, [("Do objeto", "Ação de rastreamento do câncer de mama, complementar à mamografia, por exame de sangue in house, conforme o Plano de Trabalho."),
        ("Do Plano de Trabalho", "Integra este Termo; alterações pelo art. 57."),
        ("Da natureza do exame", "Serviço laboratorial do laboratório contratado pela OSC (RDC Anvisa 978/2025), com a LiqSci interveniente; vedado custeio com recursos vinculados à tabela SUS (art. 19-T da Lei 8.080/1990)."),
        ("Das obrigações da Administração", "Integrar à APS; garantir imagem em 30 dias; designar gestor e comissão; repassar (se houver); analisar contas."),
        ("Das obrigações da OSC", "Executar o Plano; contratar laboratório habilitado; TCLE; gratuidade; LGPD; não contratar fornecedor da solução como executante; prestar contas."),
        ("Do valor", "[ R$ ___ ] — [ contrapartida / repasse ]."), ("Da vigência", "[ 12 meses ], prorrogável (art. 55)."),
        ("Do monitoramento e contas", "Arts. 58 a 72."), ("Dos dados e da pesquisa", "LGPD; CEP/CONEP para estudo (Lei 14.874/2024)."),
        ("Das sanções e rescisão", "Art. 73 e art. 42, XVI."), ("Da publicidade e foro", "Extrato publicado (art. 38); foro de [ ___ ].")])
    ass(d, ("Administração Pública", ""), ("OSC", "Representante legal")); salvar(d, "MROSC_TermoColaboracao_RosalindTest.docx")


def rede():
    d = novo(); cab(d, "TERMO DE ATUAÇÃO EM REDE (LEI Nº 13.019/2014, ART. 35-A)", "Vinculado ao Termo de Colaboração nº [ ___ ] — ação de rastreamento do câncer de mama")
    par(d, "Celebram [ OSC celebrante — CNPJ ] (CELEBRANTE) e [ OSC executante e não celebrante — CNPJ ] (EXECUTANTE):")
    clausulas(d, [("Do objeto", "Execução, pela EXECUTANTE, de acolhimento, coleta de sangue, devolutiva e navegação das usuárias, conforme o Plano de Trabalho."),
        ("Dos requisitos", "CELEBRANTE com mais de 5 anos de inscrição no CNPJ e capacidade de supervisionar a rede; EXECUTANTE com regularidade, requisitos do art. 33, sem vedações do art. 39, CNAE de saúde e responsável técnico."),
        ("Da responsabilidade", "A CELEBRANTE mantém a responsabilidade integral perante a Administração; a EXECUTANTE responde pela sua etapa."),
        ("Dos recursos", "[ valor e cronograma de transferência à EXECUTANTE ], com comprovação de despesas."),
        ("Do conflito de interesses", "A EXECUTANTE declara não fornecer nem representar a solução RosalindTest."),
        ("Dos dados", "LGPD; a EXECUTANTE como operadora."), ("Da vigência", "Vinculada ao Termo de Colaboração.")])
    ass(d, ("OSC celebrante", ""), ("OSC executante", "")); salvar(d, "MROSC_TermoAtuacaoRede_RosalindTest.docx")


def contrato_lab():
    d = novo(); cab(d, "CONTRATO DE PRESTAÇÃO DE SERVIÇOS LABORATORIAIS", "Ação de rastreamento do câncer de mama (RosalindTest) — com interveniência da titular da tecnologia")
    par(d, "CONTRATANTE: [ OSC ]. CONTRATADO: [ laboratório — CNPJ — CNES — RT ]. INTERVENIENTE ANUENTE: LiqSci Biotecnologia, CNPJ 50.433.201/0001-50.")
    clausulas(d, [("Do objeto", "Processamento de amostras de sangue pelo método in house RosalindTest, com laudo, no âmbito do Termo de Colaboração nº [ ___ ]."),
        ("Da responsabilidade técnica", "Validação, controle de qualidade, laudo e guarda de amostras pelo CONTRATADO (RDC Anvisa 978/2025)."),
        ("Do nível de serviço", "Laudo em até 20 dias; comunicação imediata de alterados; rastreabilidade; recoleta sem ônus por falha do CONTRATADO."),
        ("Da interveniência", "Licença de método e marca, suporte científico e PI pela LiqSci, que não presta o exame nem é remunerada como executante da rede. [ royalties / exclusividade ]."),
        ("Do preço", "[ R$ ___ ] por amostra com laudo; pagamento mensal por produção."), ("Dos dados", "LGPD; CEP/CONEP para estudo."),
        ("Da vigência e foro", "Vinculada ao Termo de Colaboração; foro de [ ___ ].")])
    ass(d, ("Contratante", ""), ("Contratado", ""), ("Interveniente anuente", "")); salvar(d, "Contrato_ServicoLaboratorial_LiqSci_Interveniente.docx")


def tr_cpsi():
    d = novo(); cab(d, "TERMO DE REFERÊNCIA — CPSI (LC Nº 182/2021)", "Teste de solução inovadora para ampliar o acesso ao rastreio do câncer de mama")
    h(d, "1. Problema"); par(d, "Cobertura de [ __ ]% para [ ___ ] mulheres de 40 a 74 anos; barreiras de acesso à mamografia. Busca-se solução que amplie o acesso na APS e priorize a imagem (art. 13, § 1º: descreve-se o problema, não a marca).")
    h(d, "2. Resultados esperados"); al(d, ["[ ___ ] testes com adesão mínima de [ __ ]%;", "100% das alteradas com imagem em 30 dias;", "Desempenho apurado sob protocolo CEP;", "Estadiamento dos confirmados."])
    h(d, "3. Requisitos mínimos"); al(d, ["Documentação sanitária (RDC 978/2025 ou registro Anvisa);", "Evidência prévia de desempenho;", "Coleta compatível com a APS;", "Laudo em até 20 dias; LGPD; TCLE; PGRSS."])
    h(d, "4. Julgamento (art. 13, § 4º)"); al(d, ["Potencial de resolução e economia;", "Grau de inovação;", "Viabilidade do modelo;", "Viabilidade econômica;", "Custo-benefício comparado."])
    h(d, "5. Contrato"); par(d, "CPSI de até 12 meses (+12), até R$ 1.600.000,00 (art. 14); remuneração por marcos [ 20% / 50% / 30% ]; fornecimento posterior (art. 15).")
    h(d, "6. Comissão"); par(d, "No mínimo 3 membros, incluídos 1 servidor do órgão e 1 professor de instituição pública de ensino superior na área (art. 13, § 3º).")
    ass(d, ("Responsável técnico", ""), ("Autoridade competente", "")); salvar(d, "CPSI_TermoReferencia_RosalindTest.docx")


def edital_cpsi():
    d = novo(); cab(d, "EDITAL DE LICITAÇÃO ESPECIAL Nº [ ___ ] (LC Nº 182/2021)", "Teste de solução inovadora — rastreio do câncer de mama")
    for t, txt in [("1. Objeto", "Seleção de proposta(s) de teste de solução inovadora conforme o Termo de Referência, com celebração de CPSI."),
                   ("2. Participação", "Pessoas físicas ou jurídicas, isoladas ou em consórcio; habilitação parcialmente dispensável mediante justificativa (art. 13, § 8º), mantida a regularidade com a seguridade social."),
                   ("3. Propostas", "Solução e evidência; plano de teste e protocolo CEP; preço global até R$ 1,6 mi; documentação sanitária e anuência do titular."),
                   ("4. Julgamento", "Comissão especial (art. 13, § 3º), critérios do art. 13, § 4º; possível seleção de mais de uma proposta; negociação de preço."),
                   ("5. Contrato", "CPSI (art. 14) e eventual fornecimento (art. 15)."), ("6. Recursos", "[ fonte — fora da tabela SUS ]."),
                   ("7. Disposições finais", "Confirmação por imagem garantida pela rede SUS; LGPD; Lei 14.874/2024; Lei 14.133/2021 subsidiariamente.")]:
        h(d, t); par(d, txt)
    ass(d, ("Ordenador de despesa", ""), ("Presidente da comissão", "")); salvar(d, "CPSI_Edital_RosalindTest.docx")


def contrato_cpsi():
    d = novo(); cab(d, "CONTRATO PÚBLICO PARA SOLUÇÃO INOVADORA — CPSI Nº [ ___ ]", "LC nº 182/2021, art. 14 — rastreio do câncer de mama")
    par(d, "CONTRATANTE: [ órgão ]. CONTRATADA: [ ___ ]. [ INTERVENIENTE: LiqSci Biotecnologia, CNPJ 50.433.201/0001-50 ].")
    clausulas(d, [("Do objeto", "Teste da solução conforme TR e proposta selecionada."),
        ("Das metas e da aceitação", "Relatório final com execução, navegação em 30 dias e desempenho sob protocolo CEP; regime de risco tecnológico do art. 14 [ conferir redação vigente ]."),
        ("Do preço e marcos", "[ R$ ___ ] (limite R$ 1,6 mi); marcos [ 20/50/30% ]."), ("Da matriz de riscos e PI", "Conforme anexo; PI com a titular; dados agregados à Administração."),
        ("Da vigência", "Até 12 meses, prorrogável uma vez; fornecimento posterior (art. 15)."), ("Dos dados, sanções e foro", "LGPD; Lei 14.874/2024; foro de [ ___ ].")])
    ass(d, ("Contratante", ""), ("Contratada", "")); salvar(d, "CPSI_Contrato_RosalindTest.docx")


def tcle():
    d = novo(); cab(d, "TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO (TCLE)", "Ação de rastreamento do câncer de mama — exame de sangue RosalindTest")
    for t, txt in [("1. O que é", "Exame de sangue que ajuda a priorizar quem deve fazer exame de imagem. Complementar à mamografia; não a substitui e não diagnostica."),
                   ("2. Como é feito", "Coleta de sangue do braço; laudo em cerca de 15 a 20 dias."),
                   ("3. Benefícios e limites", "Amplia o acesso ao rastreio. Pode haver falso-positivo ou falso-negativo — mantenha a mamografia de rotina. Estudos ainda iniciais."),
                   ("4. Encaminhamento", "Resultado alterado: mamografia/ultrassonografia e avaliação médica em até 30 dias, com acompanhamento."),
                   ("5. Sigilo (LGPD)", "Dados só para esta ação; laudo individual; gestão recebe números sem identificação. [ Estudo: CEP nº ___; recusa sem prejuízo. ]"),
                   ("6. Voluntariedade", "Você pode recusar ou desistir a qualquer momento sem prejuízo no SUS.")]:
        h(d, t); par(d, txt)
    par(d, "Declaro que li (ou me foi lido) este termo, entendi e concordo em participar.")
    ass(d, ("Participante", "nome / documento"), ("Profissional responsável", "nome / unidade")); salvar(d, "TCLE_RosalindTest.docx")


if __name__ == "__main__":
    for f in (nota, etp, plano, termo, rede, contrato_lab, tr_cpsi, edital_cpsi, contrato_cpsi, tcle): f()
    assert len(list(BASE.glob("*.docx"))) == 10

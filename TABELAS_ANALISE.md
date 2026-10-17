Contratos: 1001 | Obras (TipoContrato=Obra): 101 | Outliers excluídos da base 'sem outlier': 1 (pct_aditivo > 10.000%)
Campos não disponíveis na base da prefeitura: pago/empenhado, lista de termos, itens. Termos e pagamentos vêm do TCE-ES só onde há correspondência.

## 1. Totais (R$)
            grupo    n  valor_inicial          aditivo      valor_atual         delta_R$  delta_%  com_aditivo  supressoes
            Todos 1001 988,425,805.46 2,246,576,777.17 3,235,002,582.63 2,246,576,777.17   227.29          128           5
Todos sem outlier 1000 986,286,325.46   107,096,777.17 1,093,383,102.63   107,096,777.17    10.86          127           5
            Obras  101 545,988,525.59    14,161,570.88   560,150,096.47    14,161,570.88     2.59           21           1
          Serviço  707 310,051,187.52 2,227,192,659.89 2,537,243,847.41 2,227,192,659.89   718.33           82           1
         Material  151  93,180,283.40     4,568,307.11    97,748,590.51     4,568,307.11     4.90           20           3
           Outros   34  37,907,829.35       654,239.29    38,562,068.64       654,239.29     1.73            5           0

## 2. Estatística descritiva

[Todos sem outlier]
               ValorContratado       Aditivo    ValorTotal  pct_aditivo(%)
n                     1,000.00      1,000.00      1,000.00          991.00
media               986,286.33    107,096.78  1,093,383.10            8.63
mediana              50,325.00          0.00     53,066.65            0.00
desvio_padrao     4,656,289.03  1,244,039.10  5,064,340.06           32.81
min                       0.00 -1,470,524.90          0.00          -29.50
max              97,850,000.00 37,066,735.98 97,850,000.00          567.12
Somente contratos com aditivo > 0 (pct_aditivo %):
n               127.00
media            67.74
mediana          67.60
desvio_padrao    66.39
min               1.96
max             567.12

[Obras]
               ValorContratado       Aditivo    ValorTotal  pct_aditivo(%)
n                       101.00        101.00        101.00          101.00
media             5,405,826.99    140,213.57  5,546,040.56            8.00
mediana           1,478,344.72          0.00  1,539,000.00            0.00
desvio_padrao    12,529,372.34    522,915.49 12,521,583.65           23.66
min                   4,220.00 -1,470,524.90      7,820.00          -19.08
max              97,850,000.00  3,805,622.32 97,850,000.00          150.51
Somente contratos com aditivo > 0 (pct_aditivo %):
n                21.00
media            39.37
mediana          24.94
desvio_padrao    38.43
min               2.27
max             150.51

## 3. Limite legal (Obras: 25%; reforma: 50%)
Obras: 101 | com aditivo: 21 | supressão: 1 | acima do limite: 5 | entre 20% e o limite: 8

Obras com aditivo > 0 ou supressão:
NumeroContrato AnoContrato                                                  Secretaria                                                   Fornecedor  ValorContratado       Aditivo    ValorTotal  pct_aditivo  limite  excede  reforma
           223        2025                                      Secretaria de Educação KARISTEN COMÉRCIO E SERVIÇOS MECÂNICOS E ELETRICOS LTDA - ME       529,999.44    797,724.20  1,327,723.64       150.51   25.00    True    False
           325        2025 Secretaria de Transportes, Trânsito e Infraestrutura Urbana                      ILUMITERRA CONSTRUCOES E MONTAGENS LTDA     1,478,344.72  1,478,344.72  2,956,689.44       100.00   25.00    True    False
           165        2025                                 Secretaria de Meio Ambiente                      ILUMITERRA CONSTRUCOES E MONTAGENS LTDA     3,805,622.32  3,805,622.32  7,611,244.64       100.00   50.00    True     True
           293        2025                         Secretaria de Gestão e Planejamento                     INOVE ELEVADORES - PECAS E SERVICOS LTDA         4,220.00      3,600.00      7,820.00        85.31   25.00    True    False
             2        2025                                         Secretaria de Obras                                       MATIX CONSTRUTORA LTDA       904,000.00    445,915.50  1,349,915.50        49.33   25.00    True    False
            51        2025                                      Secretaria de Educação                           RENOVA CONSTRUCOES E REFORMAS LTDA     2,864,000.00  1,384,783.45  4,248,783.45        48.35   50.00   False     True
           141        2025                                      Secretaria de Educação                           RENOVA CONSTRUCOES E REFORMAS LTDA       558,208.00    246,938.61    805,146.61        44.24   50.00   False     True
            78        2025                                         Secretaria de Obras                           M.T.F.CONSTRUCOES E MONTAGENS LTDA     1,198,292.00    376,845.76  1,575,137.76        31.45   50.00   False     True
           375        2025        Secretaria de Cidadania, Direitos Humanos e Trabalho                      ILUMITERRA CONSTRUCOES E MONTAGENS LTDA       400,000.00    100,000.00    500,000.00        25.00   25.00   False    False
           309        2025                                         Secretaria de Obras                               AMF ENGENHARIA E SERVICOS LTDA       684,831.30    170,828.61    855,659.91        24.94   25.00   False    False
            25        2025                                         Secretaria de Obras        METAL EDIFICACOES E ESTRUTURAS METALICAS EIRELI - EPP     1,499,000.00    373,817.66  1,872,817.66        24.94   25.00   False    False
           474        2025                                         Secretaria de Obras        SOPE SOCIEDADE DE OBRAS E PROJETOS DE ENGENHARIA LTDA     2,403,760.64    590,881.77  2,994,642.41        24.58   25.00   False    False
            47        2026                                         Secretaria de Obras                           M.T.F.CONSTRUCOES E MONTAGENS LTDA     6,237,395.15  1,470,524.90  7,707,920.05        23.58   25.00   False    False
             5        2026 Secretaria de Transportes, Trânsito e Infraestrutura Urbana                                  VITORIALUZ CONSTRUCOES LTDA     1,159,962.62    216,753.22  1,376,715.84        18.69   25.00   False    False
            89        2025                                         Secretaria de Obras                                      APPOGI CONSTRUTORA LTDA     6,557,852.04  1,137,523.21  7,695,375.25        17.35   25.00   False    False
             3        2025                                         Secretaria de Obras        SOPE SOCIEDADE DE OBRAS E PROJETOS DE ENGENHARIA LTDA     9,946,039.66  1,698,485.85 11,644,525.51        17.08   25.00   False    False
           432        2025                                         Secretaria de Obras                           L. A. MONJARDIM CONSTRUTORA EIRELI       959,899.86    137,717.84  1,097,617.70        14.35   25.00   False    False
           230        2025                                         Secretaria de Obras                           RENOVA CONSTRUCOES E REFORMAS LTDA     3,500,000.00    437,405.96  3,937,405.96        12.50   25.00   False    False
           282        2025                                         Secretaria de Obras                         ENGESAN ENGENHARIA E SANEAMENTO LTDA     7,146,686.59    638,931.42  7,785,618.01         8.94   25.00   False    False
           274        2025                                       Secretaria de Fazenda                     CITY CONNECT SOLUCOES EM TECNOLOGIA LTDA     2,541,796.93     84,258.99  2,626,055.92         3.31   25.00   False    False
            14        2026                                         Secretaria de Obras                           M.T.F.CONSTRUCOES E MONTAGENS LTDA     1,550,000.00     35,191.79  1,585,191.79         2.27   25.00   False    False
            39        2026                                         Secretaria de Obras                           M.T.F.CONSTRUCOES E MONTAGENS LTDA     7,707,920.05 -1,470,524.90  6,237,395.15       -19.08   25.00   False    False

Não-obras com pct_aditivo > 25%: 64 (sem outlier: 63); soma dos aditivos sem outlier: 89,954,162.89

## 4. Discrepâncias matemáticas
ValorContratado + Aditivo != ValorTotal (tol. R$ 0,01): 0
ValorContratado <= 0: 9 | Aditivo > ValorContratado: 39

PMV x TCE-ES (contratos casados: 615; obras: 19):
  valor inicial igual: 615
  aditivo>0 na PMV e sem termo no TCE: 7 | termo aditivo no TCE e aditivo<=0 na PMV: 4
  contratos com termo no TCE: 49 (aditivos: 43, apostilamentos: 20, prorrogações: 21, acréscimos/supressões: 9, reequilíbrio: 6)
  contratos com pagamento no TCE: 222 | total pago: 37,785,041.31

## 5. Licitação (PNCP) x contrato
Contratos com licitação no PNCP: 792 (obras: 62) | com valor homologado: 718 (obras: 34)

[Todos] (estimado-homologado)/estimado (%)
n                718.00
media             10.50
mediana            0.00
desvio_padrao     22.78
min             -125.00
max               93.01
ValorContratado > homologado: 3 | ValorTotal > homologado: 32

[Obras] (estimado-homologado)/estimado (%)
n               34.00
media           12.58
mediana          4.62
desvio_padrao   14.03
min              0.00
max             43.72
ValorContratado > homologado: 0 | ValorTotal > homologado: 4

Por modalidade (sem outlier):
                             n        inicial       aditivo  com_aditivo  pct_mediana  aditivo/inicial_%
modalidadeNome                                                                                          
Concorrência - Eletrônica   28 158,483,746.92    452,959.01            3         0.00               0.29
Dispensa                    94  54,343,376.06  2,625,731.66           15         0.00               4.83
Inexigibilidade            431  47,315,925.43 13,311,121.98           14         0.00              28.13
Pregão - Eletrônico        236 237,475,757.03 10,999,943.71           34         0.00               4.63
Pregão - Presencial          2     588,000.00          0.00            0         0.00               0.00

Obras por modalidade:
                            n        inicial      aditivo  com_aditivo  aditivo/inicial_%
modalidadeNome                                                                           
Concorrência - Eletrônica  26 157,845,746.92   452,959.01            3               0.29
Dispensa                    9  14,360,163.64   594,481.77            2               4.14
Inexigibilidade            13   1,172,148.35         0.00            0               0.00
Pregão - Eletrônico        14  48,973,164.22 4,920,099.74            4              10.05

## 6. Por secretaria (sem outlier, top 10 por aditivo)
                                                               n        inicial       aditivo  com_aditivo  aditivo/inicial_%
Secretaria                                                                                                                   
Secretaria de Educação                                        43 123,790,944.68 47,284,794.27           17              38.20
FUNDO MUNICIPAL DE SAÚDE DE VITORIA                          107 151,642,825.51 30,165,127.78           30              19.89
Secretaria de Obras                                           59 352,922,215.29  6,554,812.01           14               1.86
Secretaria de Meio Ambiente                                   15  19,655,922.26  5,115,932.81            4              26.03
Secretaria de Assistência Social                              18   6,582,823.31  4,198,038.06            5              63.77
Secretaria de Esportes e Lazer                                32   5,885,148.27  2,507,504.07            7              42.61
Secretaria de Gestão e Planejamento                           50  19,126,072.86  1,906,334.21            9               9.97
Central de Serviços                                           27  66,523,009.43  1,828,691.90            2               2.75
Secretaria de Transportes, Trânsito e Infraestrutura Urbana   17  36,663,644.77  1,706,530.48            3               4.65
Secretaria de Cultura                                        350  20,330,328.46  1,376,357.91            5               6.77

Obras por secretaria:
                                                              n        inicial      aditivo  com_aditivo  aditivo/inicial_%
Secretaria                                                                                                                 
Secretaria de Obras                                          40 350,586,908.03 6,043,545.37           12               1.72
Central de Serviços                                           9  50,764,588.00         0.00            0               0.00
FUNDO MUNICIPAL DE SAÚDE DE VITORIA                           4  43,358,630.39         0.00            0               0.00
Secretaria de Educação                                        8  39,413,163.52 2,429,446.26            3               6.16
Secretaria de Transportes, Trânsito e Infraestrutura Urbana   5  33,463,993.05 1,695,097.94            2               5.07
Fundo Municipal de Educação                                   5   8,213,500.00         0.00            0               0.00
Secretaria de Fazenda                                         4   5,625,099.05    84,258.99            1               1.50
Secretaria de Meio Ambiente                                   1   3,805,622.32 3,805,622.32            1             100.00
Secretaria de Desenvolvimento da Cidade e Habitação          14   3,544,020.10         0.00            0               0.00
Secretaria de Gestão e Planejamento                           2   3,532,220.00     3,600.00            1               0.10
Secretaria de Cultura                                         2   1,109,000.00         0.00            0               0.00
Secretaria de Esportes e Lazer                                1     800,000.00         0.00            0               0.00
Fundo Municipal de Cultura                                    1     620,000.00         0.00            0               0.00
Secretaria de Segurança Urbana                                2     514,000.00         0.00            0               0.00
Secretaria de Cidadania, Direitos Humanos e Trabalho          1     400,000.00   100,000.00            1              25.00
Fundo Municipal de Assistência Social                         1     125,200.00         0.00            0               0.00
FUNDO MUNICIPAL DE MEIO AMBIENTE                              1     112,581.13         0.00            0               0.00
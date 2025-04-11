# Sistema de Cálculo de Impostos

## **Descrição Geral**
Este sistema foi desenvolvido em VBA no Excel com o objetivo principal de **calcular impostos** de empresas enquadradas nos regimes tributários de **lucro real** e **lucro presumido**. O cálculo inclui os seguintes impostos:

- **IRPJ** (em desenvolvimento)
- **CSLL** (em desenvolvimento)
- **PIS**
- **COFINS**

---

## **Funcionalidades**

### 1. **Cálculo de Impostos**
   - Realiza o cálculo automático dos impostos **PIS**, **COFINS**, **IRPJ** e **CSLL**.
   - Suporte a empresas de lucro real e lucro presumido.

### 2. **Importação de Notas Fiscais**
   - Importa **notas fiscais de serviço** no formato XML emitidas pela prefeitura de Belo Horizonte (BH).
   - Simplifica o processo de preenchimento de dados necessários para os cálculos.

### 3. **Verificação de Numeração de Notas**
   - Verifica a sequência de numeração das notas fiscais importadas.
   - Identifica descontinuidades, indicando possíveis notas faltantes.

---

## **Requisitos Técnicos**

- **Ambiente**: Microsoft Excel com suporte ao VBA.
- **Versão do MSXML**: 6.0.
- **Notas Fiscais Suportadas**: Apenas notas fiscais de serviço emitidas pela prefeitura de Belo Horizonte (BH).

---

## **Status de Desenvolvimento**

- **Cálculo de IRPJ e CSLL**: Funcionalidade em desenvolvimento.
- **Importação de Notas**: Totalmente funcional.
- **Verificação de Numeração**: Totalmente funcional.

---

## **Como Utilizar**
1. Abra o Excel e habilite macros.
2. Certifique-se de que o MSXML 6.0 está instalado e ativo no seu sistema.
3. Importe os arquivos XML das notas fiscais pela funcionalidade disponibilizada.
4. Verifique a sequência de numeração para garantir a consistência dos dados.
5. Utilize os dados das notas para o cálculo dos impostos.

---

## **Próximos Passos**
- Concluir o desenvolvimento dos cálculos de **IRPJ** e **CSLL**.
- Expandir o suporte para notas fiscais de outros municípios (opcional).
- Otimizar a interface e funcionalidades para maior usabilidade.

---

## **Observações**
Este sistema foi projetado para atender às necessidades específicas de cálculo tributário, proporcionando eficiência e precisão ao processo. Embora a importação de notas fiscais seja uma funcionalidade útil, ela é secundária ao objetivo principal de realizar os cálculos tributários.

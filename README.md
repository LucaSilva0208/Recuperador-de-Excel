# 📊 Recuperador de Excel

Aplicação de resgate de dados (Data Rescue) para planilhas `.xlsx` corrompidas ou protegidas.

### 🛠️ Tecnologias e Arquitetura
- **Backend:** Node.js + Express.
- **Processamento:** ExcelJS com **Streaming de Leitura/Escrita** (otimizado para baixo consumo de memória).
- **Frontend:** HTML5, CSS3 e JavaScript Vanilla (Leve e responsivo).
- **Hospedagem:** Render (PaaS).

### ⚙️ Diferenciais de Engenharia (Visão de QA/Analista)
- **Gerenciamento de Memória:** O sistema não carrega o arquivo inteiro na RAM. Ele lê e escreve linha por linha, permitindo processar grandes volumes em instâncias limitadas.
- **Sanitização Automatizada:** Remoção de caracteres de controle ASCII que impedem a abertura do XML pelo Microsoft Excel.
- **Política Zero-Persistence:** Arquivos originais e recuperados são excluídos automaticamente para garantir a segurança e economia de disco.

### 📋 Requisitos e Regras de Negócio
- Limite de 25MB por arquivo.
- Suporte exclusivo para extensão `.xlsx`.
- Bypass automático de proteção de planilha/pasta de trabalho.

---
Desenvolvido por **Lucas Silva** | Foco em Qualidade e Análise de Sistemas.

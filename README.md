# 📝 Buscador de Palavras-chave em Planilhas Excel

## 📌 1. O que o programa faz

O **Buscador de Palavras-chave em Excel** permite procurar palavras ou expressões dentro de planilhas, com suporte para diferentes modos de busca:

- **Busca exata**  
- **Busca por padrão (regex)**  
- **Busca por similaridade (fuzzy)** com ajuste de porcentagem  
- **Busca radical (stem)** 

Ele localiza todas as ocorrências e exibe:

- Número total de ocorrências  
- Resumo de quantas vezes cada palavra apareceu  
- Linhas e colunas onde encontrou  
- Trecho do conteúdo original  

Também permite **salvar os resultados** em um novo Excel com todas as colunas originais.

---

## 🚀 2. Como usar

### **Passo 1 – Abrir o programa**

Abra o arquivo `Busca por Palavra-chave.exe`.

![Tela inicial](/imagens/tela_inicial.png)

---

### **Passo 2 – Selecionar o arquivo**

Clique em **Selecionar Arquivo** e escolha seu Excel.

![Seleção de arquivo](/imagens/selecionar_arquivo.png)

---

### **Passo 3 – Escolher a aba**

No campo “Aba”, escolha a planilha onde deseja buscar.

![Seleção de aba](/imagens/selecionar_aba.png)

---

### **Passo 4 – Inserir palavras-chave**

Digite as palavras separadas por vírgula:

```

arquivamento, desarquivamento, arquivar

```

![Inserindo palavras](/imagens/inserir_palavras.png)

---

### **Passo 5 – Configurar opções**

- **Modo de busca**:
  - `exato` → Igualdade literal.
  - `padrão` → Padrões avançados.
  - `similaridade` → Palavras semelhantes (ajuste no slider *% de Similaridade*).
  - `redical` → Radicais das palavras.

- **Colunas específicas (opcional)**: marque a opção, clique em **Ver Colunas** e selecione.

![Opções avançadas](/imagens/opcoes_busca.png)

---

### **Passo 6 – Executar a busca**

Clique em **Buscar** e aguarde.

![Progresso da busca](/imagens/busca_em_andamento.png)

---

### **Passo 7 – Resultados**

Veja o total de ocorrências, resumo por palavra e detalhes de cada linha.

![Resultados](/imagens/resultados_busca.png)

---

### **Passo 8 – Salvar resultados**

Clique em **Salvar Resultados** e escolha onde gravar.

![Salvar resultados](/imagens/salvar_resultados.png)

---

## 📊 3. Exemplos de uso

| Cenário | Configuração sugerida |
|---------|-----------------------|
| Palavras exatas | `exato` |
| Prefixos/sufixos | `padrão` |
| Palavras parecidas | `similaridade` + Similaridade 80 |
| Variações de conjugação | `radical` |

---

## 💡 4. Dicas

- Limitar colunas acelera buscas grandes.  
- No **similaridade**, valores muito baixos (<70) aumentam falsos positivos.  
- **padrão** exige conhecimento em expressões regulares.  
- Arquivos muito grandes podem levar mais tempo.

---

## ⚙️ 5. Requisitos

- Windows 10 ou superior.  
- Não precisa ter Excel instalado.  
- Evite planilhas com mais de **200 mil linhas** sem filtro de coluna.

---

# LiteFlow 🌊

O **LiteFlow** é um motor de gerenciamento de evidências e geração de relatórios focado em altíssima performance e baixíssimo consumo de memória. Nascido para complementar o ecossistema LiteTools (atuando em conjunto com ferramentas de captura como o LiteShot), ele organiza suas capturas em uma esteira visual, permitindo edição avançada, anotações detalhadas, reordenação e exportação final estruturada para Word (.docx) ou PDF.

Desenvolvido em C# (.NET 10) e Windows Forms, sua arquitetura foi exaustivamente desenhada para testes de estresse (suportando fluxos de 100+ imagens pesadas) utilizando I/O de disco direto e virtualização de UI, garantindo que a máquina nunca trave por falta de memória RAM durante a montagem de evidências críticas de QA.

<br>

## 🔒 Segurança e Privacidade

> [!IMPORTANT]
> **O LiteFlow é uma ferramenta 100% offline.**
>
> * Não realiza chamadas de rede.
> * Não possui telemetria.
> * Não coleta métricas de uso.
> * Não envia dados ou relatórios para a nuvem.
> * Não possui mecanismo de auto-update.
>
> Todo o processamento, geração de PDFs e manipulação de documentos Word ocorre localmente na sua máquina. O código é open source para permitir auditoria e total transparência.

<br>

## 🎯 Non-Goals (O que NÃO fazemos)

> [!WARNING]
> **O LiteFlow não tem como objetivo:**
>
> * Substituir suítes completas na nuvem (como Jira ou Xray).
> * Fazer backup ou sincronização de projetos em servidores remotos.
> * Coletar dados dos cenários de teste gerados (telemetria zero).
>
> *O foco é estritamente a produtividade local, organização rápida de imagens e montagem limpa e automatizada de documentos finais.*

<br>

## 📥 Download

Você não precisa baixar o código-fonte para usar! Baixe a versão pronta para uso diretamente na página de Releases do GitHub:

👉 [**Baixar LiteShot (Versão Mais Recente)**](https://github.com/eugenio122/LiteFlow/releases/latest)

<br>

## ✨ Funcionalidades e Engenharia

* **Arquitetura "Zero Leak" (Disk-Backed State):** Mantém o consumo de RAM estabilizado de forma agressiva (suportando dezenas de prints). O sistema salva as imagens em altíssima qualidade no armazenamento local (SSD/HDD) no momento da captura. Apenas as miniaturas ultracomprimidas e as 3 imagens mais recentes ativas permanecem na memória RAM. O restante do histórico é gerido por Lazy Load.

* **Captura O(1) (Fire and Forget):** Injeção de imagens assíncrona. O host não congela ao enviar uma nova captura, garantindo que adicionar a 1ª ou a 100ª imagem ao relatório demore exatamente a mesma fração de milissegundos.

* **Editor Nativo Integrado:** Edite a evidência no momento em que a seleciona. Inclui Crop (recorte destrutivo otimizado), Caneta, Marcador, Linha, Seta, Formas e Texto. Ações baseadas em pilhas de Undo/Redo sem vazamento de memória.

* **Exportação Otimizada (Direct I/O):** Exporte históricos imensos para Word ou PDF utilizando Templates base configuráveis. O motor de exportação lê os arquivos em streaming direto do disco físico para evitar gargalos (RAM Spikes) comuns em manipulação pesada via GDI+.

* **Auto-Save Inteligente e Resiliência:** Gravação em background via JSON estruturado. Mesmo em caso de queda de energia no passo 90 do seu teste, todo o trabalho visual e anotações estarão salvos.

* **Organização e Automação:** Arraste e solte para reordenar passos do teste. Adicione anotações de evidência, posicione texto acima ou abaixo da imagem, e utilize a tag "Apenas Evidência" (👁️) para integrações futuras com motores de geração (LiteJson).

* **Multilíngue (i18n):** Suporte nativo para Português, Inglês, Espanhol, Francês, Alemão e Italiano.

<br>

## 🚀 Como usar

1. Execute o LiteFlow (Modo Standalone) ou abra-o embarcado no ecossistema LiteTools.
2. Inicie suas capturas de tela (via LiteShot, colando Ctrl+V da área de transferência ou criando telas em branco).
3. Selecione as imagens na esteira inferior (Ribbon) para editá-las ou adicionar anotações (descrição do passo).
4. Utilize o arrastar-e-soltar para organizar a cronologia do seu cenário de teste.
5. No painel direito, ajuste as propriedades de layout e clique em Exportar WORD ou Exportar PDF.

<br>

## 🛠️ Como compilar (Para Desenvolvedores)

Este projeto usa o .NET 10 e Windows Forms. Você pode compilá-lo para gerar o executável/DLL portátil da seguinte forma:

**Build Portátil (Arquivo único, roda em qualquer PC):**

```bash
dotnet publish -c Release -r win-x64 --self-contained true
```

/// <reference lib="dom"/>

import * as Puppeteer from 'npm:puppeteer@21.2.1';
import * as ChartJS from 'npm:chart.js@4.4.0';

import Xlsx from "npm:xlsx@0.18.5";

import path from 'node:path';
import fs from 'node:fs';

async function programa() {
  const caminho_da_planilha = Deno.args[0] || prompt("Digite o caminho do arquivo de planilha:");
  const caminho_de_saída = Deno.args[1] || prompt("Digite o caminho de saída dos gráficos:");

  if (!caminho_da_planilha || !caminho_de_saída) return;

  const analisador = new Analisador(caminho_da_planilha, caminho_de_saída);

  await Promise.all([
    analisador.tratar_dados(),
    analisador.exibir_análise(),
    analisador.renderizar_gráficos()
  ])

  Deno.exit(0);
}

class Analisador {
  private readonly caminho_da_planilha: string;
  private readonly caminho_de_saída: string;

  private readonly lista_de_informações: Array<Informação>;

  private readonly contagem_de_categorias: Contagem;
  private readonly contagem_de_palavras: Contagem;
  
  private readonly interações: Interação;

  constructor(caminho_da_planilha: string, caminho_de_saída: string) {
    if (!fs.existsSync(caminho_de_saída)) fs.mkdirSync(caminho_de_saída);

    this.caminho_da_planilha = caminho_da_planilha;
    this.caminho_de_saída = caminho_de_saída;

    const pasta_de_trabalho = Xlsx.readFile(this.caminho_da_planilha);
    const nome_da_planilha_principal = pasta_de_trabalho.SheetNames[0];
    const planilha_principal = pasta_de_trabalho.Sheets[nome_da_planilha_principal];

    this.lista_de_informações = Xlsx.utils.sheet_to_json(planilha_principal);

    this.contagem_de_categorias = {
      [Analisador.CATEGORIAS.CARREIRA]: 0,
      [Analisador.CATEGORIAS.FINANÇAS]: 0,
      [Analisador.CATEGORIAS.EDUCAÇÃO]: 0,
      [Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS]: 0,
    }

    this.contagem_de_palavras = {};
    
    this.interações = {};
  }

  public tratar_dados() {
    for (const informação of this.lista_de_informações) {
      for (const pergunta in informação) {
        const resposta = String(informação[pergunta]);
        const elementos = resposta.match(Analisador.EXPRESSÃO_DE_LISTAS)?.map((item) => item.slice(0, -1));
        const estáDuplicada = pergunta.match(Analisador.EXPRESSÃO_DE_DUPLICADOS);

        if (estáDuplicada) {
          this.unir_pergunta_duplicada(pergunta, informação);
        } else {
          this.registrar_informação(pergunta, resposta, elementos);
        }

        if (pergunta === "Escreva algumas linhas sobre sua história e seus sonhos de vida.") {
          const palavras_da_resposta = resposta.toLowerCase().match(Analisador.EXPRESSÃO_DE_PALAVRAS) || [];

          palavras_da_resposta.forEach((palavra) => {
            const categoria_da_palavra = Analisador.PALAVRAS_CHAVE[palavra];

            if (categoria_da_palavra) {
              if (!this.contagem_de_categorias[categoria_da_palavra]) this.contagem_de_categorias[categoria_da_palavra] = 0;
              if (!this.contagem_de_palavras[categoria_da_palavra]) this.contagem_de_palavras[categoria_da_palavra] = 0;

              this.contagem_de_categorias[categoria_da_palavra]++;
              this.contagem_de_palavras[categoria_da_palavra]++;
            }
          });
        }

        if (pergunta === "Qual empresa que você está contratado agora?") {
          const palavras_da_resposta_sem_paradas = this.remover_paradas(resposta);

          palavras_da_resposta_sem_paradas.forEach((palavra) => {
            if (!this.contagem_de_palavras[palavra]) this.contagem_de_palavras[palavra] = 0;
            
            this.contagem_de_palavras[palavra]++;
          });
        }
      }
    }

    for (const categoria in this.contagem_de_categorias) {
      if (!this.interações[Analisador.PERGUNTA_PARA_CATEGORIAS]) this.interações[Analisador.PERGUNTA_PARA_CATEGORIAS] = {};
      this.interações[Analisador.PERGUNTA_PARA_CATEGORIAS][categoria] = this.contagem_de_categorias[categoria];
    }

    for (let pergunta in this.interações) {
      const cópia = this.interações[pergunta];
      delete this.interações[pergunta];

      pergunta = pergunta.replace('V+AZ:CP', 'V').replaceAll(' ?', '?');
      this.interações[pergunta] = cópia;
    }
  }

  public async renderizar_gráficos() {  
    const configurações_em_json: Array<string> = [];
    let contador_de_perguntas = 1;

    const navegador = await Puppeteer.launch({headless: 'new'});
    const página = await navegador.newPage();

    await página.goto('about:blank');

    await página.evaluate(() => {
      document.body.style.display = 'flex';
      document.body.style.flexDirection = 'column';
      document.body.style.alignItems = 'center';
      document.body.style.justifyContent = 'center';
    });

    await página.evaluate(async () => {
      const script = document.createElement('script');
      
      script.innerHTML = await fetch('https://cdn.jsdelivr.net/npm/chart.js').then((resposta) => resposta.text());
      script.async = true;
      
      document.head.append(script);
    });

    for (const pergunta in this.interações) {
      const respostas: Array<string> = [];
      const contagem: Array<number> = [];

      for (const resposta in this.interações[pergunta]) {
        respostas.push(resposta);
        contagem.push(this.interações[pergunta][resposta]);
      }

      const éContável = contagem.reduce((acumulador, elemento) => acumulador + elemento, 0) !== contagem.length;
      if (!éContável) continue;

      const resultado = await página.evaluate((pergunta, respostas, contagem, contador_de_perguntas) => {
        return new Promise((resolve) => {
          const canvas = document.createElement('canvas');

          canvas.width = 1300;
          canvas.height = 850;
          canvas.id = String(contador_de_perguntas);

          document.body.append(canvas);

          const contexto = canvas.getContext('2d');
          if (!contexto) return;

          const configuração: ChartJS.ChartConfiguration<"pie", number[], string> = {
            type: 'pie',
            data: {
              labels: respostas,
              datasets: [{
                data: contagem
              }]
            },
            options: {
              radius: 250,
              responsive: false,
              plugins: {
                title: {
                  display: true,
                  text: pergunta,
                  font: {
                    size: 35
                  }
                },
                legend: {
                  display: true,
                  position: 'top',
                  labels: {
                    font: {
                      size: 15
                    }
                  }
                }
              },
              animation: {
                duration: 0,
                onComplete: () => {
                  const resultado: Resultado = {
                    imagem_em_base64: canvas.toDataURL('image/png'),
                    configuração_em_json: JSON.stringify(configuração)
                  }

                  resolve(resultado);
                }
              }
            }
          }

          // @ts-ignore:
          new Chart(contexto, configuração);
        })
      }, pergunta, respostas, contagem, contador_de_perguntas) as Resultado;

      if (resultado) {
        const dados = String(resultado.imagem_em_base64).split(',')[1];
        const bytes = atob(dados);
        const binário = new Uint8Array(bytes.length);

        for (let i = 0; i < bytes.length; i++) binário[i] = bytes.charCodeAt(i);

        const nome_do_arquivo = path.join(this.caminho_de_saída, `Gráfico ${contador_de_perguntas}.png`);

        console.log('\x1b[32mSalvando\x1b[0m', nome_do_arquivo);
        fs.writeFileSync(nome_do_arquivo, binário);

        contador_de_perguntas++;
        configurações_em_json.push(resultado.configuração_em_json);
      }
    }

    await página.evaluate((configurações_em_json) => {
      document.body.innerHTML += `
        <script>
          const configurações_em_json = [${configurações_em_json}];

          configurações_em_json.forEach((configuração, índice) => {
            const id = String(índice + 1);
            const canvas = document.getElementById(id);
            const contexto = canvas.getContext('2d');
            
            new Chart(contexto, configuração);
          });
        </script>
      `;
    }, configurações_em_json);

    const código_html = await página.evaluate(() => {
      return document.querySelector('*')?.outerHTML;
    });

    if (código_html) {
      const nome_do_arquivo = path.join(this.caminho_de_saída, 'Gráficos.html');

      console.log('\x1b[32mSalvando\x1b[0m', nome_do_arquivo);
      fs.writeFileSync(nome_do_arquivo, código_html);
    }

    await navegador.close();
  }

  public exibir_análise() {
    for (const categoria in this.contagem_de_categorias) {
      console.log(`Total de palavras-chave na categoria ${categoria}: ${this.contagem_de_categorias[categoria]}`);
    }

    console.log("Contagem global de palavras:", this.contagem_de_palavras);
    console.log("Contagem das perguntas e respostas:", this.interações);
    console.log("Lista de informacoes:", this.lista_de_informações);
  }

  private unir_pergunta_duplicada(pergunta: string, informacao: Informação) {
    const perguntaCompleta = pergunta.replace(Analisador.EXPRESSÃO_DE_DUPLICADOS, "");
    if (informacao[pergunta]) informacao[perguntaCompleta] = informacao[pergunta];
  }

  private registrar_informação(pergunta: string, resposta: string, elementos: Array<string> | undefined) {
    if (!this.interações[pergunta]) this.interações[pergunta] = {};

    if (elementos) {
      elementos.forEach((elemento) => {
        const resposta_contada = this.interações[pergunta][elemento];
        if (!resposta_contada) this.interações[pergunta][elemento] = 0;

        this.interações[pergunta][elemento]++;
      });
    }
    
    else {
      const resposta_contada = this.interações[pergunta][resposta];
      if (!resposta_contada) this.interações[pergunta][resposta] = 0;

      this.interações[pergunta][resposta]++;
    }
  }

  private remover_paradas(resposta: string) {
    const palavras = resposta.toLowerCase().match(Analisador.EXPRESSÃO_DE_PALAVRAS) || [];
    const palavras_sem_paradas: Array<string> = [];

    palavras.forEach((palavra) => {
      if (!Analisador.PARADAS.includes(palavra)) palavras_sem_paradas.push(palavra);
    });
  
    return palavras_sem_paradas;
  }

  private static CATEGORIAS: Categorias = {
    CARREIRA: "Carreira (pessoas que querem garantir um emprego/carreira na área)",
    FINANÇAS: "Finanças (pessoas que buscam uma vida financeira estável com o curso)",
    EDUCAÇÃO: "Educação (pessoas que querem buscar conhecimento/ou a conclusão do curso)",
    RELAÇÕES_PESSOAIS: "Relações Pessoais (pessoas que começaram a estudar na área por amigos ou parentes e que querem sustentá-los)"
  }

  private static PALAVRAS_CHAVE: Chaves = {
    "trabalho": Analisador.CATEGORIAS.CARREIRA,
    "emprego": Analisador.CATEGORIAS.CARREIRA,
    "profissionalizar": Analisador.CATEGORIAS.CARREIRA,
    "trabalhar": Analisador.CATEGORIAS.CARREIRA,
    "área": Analisador.CATEGORIAS.CARREIRA,
    "conhecimento": Analisador.CATEGORIAS.EDUCAÇÃO,
    "formar": Analisador.CATEGORIAS.EDUCAÇÃO,
    "graduar": Analisador.CATEGORIAS.EDUCAÇÃO,
    "graduado": Analisador.CATEGORIAS.EDUCAÇÃO,
    "especializar": Analisador.CATEGORIAS.EDUCAÇÃO,
    "especialização": Analisador.CATEGORIAS.EDUCAÇÃO,
    "estudar": Analisador.CATEGORIAS.EDUCAÇÃO,
    "curso": Analisador.CATEGORIAS.EDUCAÇÃO,
    "diploma": Analisador.CATEGORIAS.EDUCAÇÃO,
    "família": Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS,
    "mãe": Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS,
    "pai": Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS,
    "amigos": Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS,
    "intercâmbio": Analisador.CATEGORIAS.RELAÇÕES_PESSOAIS,
    "estabilidade": Analisador.CATEGORIAS.FINANÇAS,
    "dinheiro": Analisador.CATEGORIAS.FINANÇAS,
    "sustentar": Analisador.CATEGORIAS.FINANÇAS,
  };

  private static EXPRESSÃO_DE_DUPLICADOS = /\d+$/g;
  private static EXPRESSÃO_DE_PALAVRAS = /\S+/g;
  private static EXPRESSÃO_DE_LISTAS = /([^;\s]+);/g;

  public static PARADAS = ["a", "o", "em", "de", "para", "com", "que", "você", "está"];

  public static PERGUNTA_PARA_CATEGORIAS = "Distribuição de alunos que responderam o questionário do perfil socioeconômico";
}

interface Informação {
  [informação: string]: string | number
}

interface Contagem {
  [palavra: string]: number;
}

interface Interação {
  [pergunta: string]: {
    [resposta: string]: number
  }
}

interface Categorias {
  [nome_da_categoria: string]: string
}

interface Chaves {
  [palavra_chave: string]: string
}

interface Resultado {
  imagem_em_base64: string,
  configuração_em_json: string
}

programa();
import dotenv from "dotenv";
dotenv.config();

import { getAuth, infoAPI } from "./getAuth.js";
import ExcelJS from "exceljs";
import nodemailer from "nodemailer"

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

var date = new Date();
var today = `${String(date.getFullYear())}-${String(
  date.getMonth() + 1
).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
let qtdItens;

async function getSinistros() {
  try {
    const JSONdata = await getAuth();
    const token = JSONdata.data.token;
    let page = 1;
    let allResults = [];
    let nextPage;

    do {
      const response = await fetch(
        `${infoAPI.url}/sinistros/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataAviso",
            dataInicial: "2023-01-01",
            dataFinal: today,
          }),
        }
      );

      if (!response.ok) {
        console.log(`Erro na requisição: ${response.status}`);
        break;
      }

      const result = await response.json();
      nextPage = result.links.next;
      const sinistrosData = result.data;

      for (const sinistros of sinistrosData) {
        allResults.push(sinistros);
      }

      console.log(`Iterando... ${allResults.length}`);

      page++;

      await delay(1000);
    } while (nextPage !== null);

    console.log(`Quantidade total de sinistros: ${allResults.length}`);
    qtdItens = allResults.length

    return allResults;
  } catch (err) {
    console.log(`Erro: ${err}`);
  }
}

async function saveToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("tabelaSinistros");

  const columns = [
    { header: "Sinistro ID", key: "sinistroId", width: 15 },
    { header: "Valor Indenizado", key: "valorIndenizado", width: 20 },
    { header: "Data do Aviso", key: "dataAviso", width: 20 },
    { header: "Data do Sinistro", key: "dataSinistro", width: 20 },
    { header: "Data da Vistoria", key: "dataVistoria", width: 10 },
    { header: "Data do Pagamento", key: "dataPagamento", width: 20 },
    {
      header: "Data Autorização de Reparos",
      key: "dataAutorizacaoReparos",
      width: 15,
    },
    { header: "Data Envio da NF", key: "dataEnvioNF", width: 20 },
    { header: "Data Documentação", key: "dataDocumentacao", width: 20 },
    { header: "Corretora", key: "corretoresNome", width: 20 },
    { header: "Seguradora", key: "companhiaNome", width: 20 },
    { header: "Segurado", key: "seguradoNome", width: 20 },
    { header: "CPF/CNPJ", key: "seguradoCpf_cnpj", width: 20 },
    { header: "Status", key: "statusSinistroNome", width: 20 },
    { header: "Ramo", key: "ramoNome", width: 20 },
    { header: "Produtor de Repasse", key: "repassesNome", width: 20 },
    { header: "Tipo do Sinistro", key: "tipoNome", width: 20 },
  ];

  worksheet.columns = columns;

  data.forEach((item) => {
    worksheet.addRow({
      sinistroId: item.sinistroId || "",
      valorIndenizado: item.valorIndenizado || "",
      dataAviso: item.dataAviso || "",
      dataSinistro: item.dataSinistro || "",
      dataVistoria: item.dataVistoria || "",
      dataPagamento: item.dataPagamento || "",
      dataAutorizacaoReparos: item.dataAutorizacaoReparos || "",
      dataEnvioNF: item.dataEnvioNF || "",
      dataDocumentacao: item.dataDocumentacao || "",
      corretoresNome: item.proposta.corretores[0].nome || "",
      companhiaNome: item.companhia.nome || "",
      seguradoNome: item.proposta.segurado.nome || "",
      seguradoCpf_cnpj: item.proposta.segurado.cpf_cnpj || "",
      statusSinistroNome: item.statusSinistro?.nome || "",
      ramoNome: item.proposta.ramo.nome || "",
      repassesNome: item.proposta.repasses[0].produtor.nome || "",
      tipoNome: item.tipo.nome || "",
    });
  });

  workbook.xlsx
    .writeFile("tabelaSinistros.xlsx")
    .then(() => {
      console.log("Arquivo Excel gerado com sucesso!");
    })
    .catch((err) => {
      console.error("Erro ao gerar o arquivo Excel:", err);
    });
}

async function main() {
  const sinistrosData = await getSinistros();
  await saveToExcel(sinistrosData);

  const email = process.env.MAIL_EMAIL;
  const password = process.env.MAIL_PASSWORD;
  var dateAndHour = `${String(date.getFullYear())}-${String(
    date.getMonth() + 1
  ).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")} às ${String(
    date.getHours()
  )}:${String(date.getMinutes())}:${String(date.getSeconds())}`;

  let transporter = nodemailer.createTransport({
    host: "smtp.hostinger.com",
    port: 465,
    auth: {
      user: email,
      pass: password,
    },
  });

  transporter.sendMail({
    from: email,
    to: email,
    subject: "A Tabela de Sinistros foi Atualizada!",
    html: `<h1 style="color: green">Sinistros Atualizados com Sucesso!</h1>
    <p>Segue os dados:</p>
    <ul>
    <li>Quantidade de Sinistros: ${qtdItens}</li>
    <li>Atualizado dia ${dateAndHour}</li>
    </ul>
    `,
  })
  .then(() => {
    console.log('E-mail SINISTRO Enviado!')
  })
  .catch((err) => {
    console.log(`Erro ao enviar o e-mail ${err}`)
  })
}

main();

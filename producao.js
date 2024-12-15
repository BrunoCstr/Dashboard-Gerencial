import dotenv from "dotenv";
dotenv.config();

import { getAuth, infoAPI } from "./getAuth.js";
import ExcelJS from "exceljs";
import nodemailer from "nodemailer";

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

var date = new Date();
var today = `${String(date.getFullYear())}-${String(
  date.getMonth() + 1
).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
let qtdItens;

async function getProducao() {
  try {
    const JSONdata = await getAuth();
    const token = JSONdata.data.token;
    let page = 1;
    let allResults = [];
    let nextPage;

    do {
      const response = await fetch(
        `${infoAPI.url}/producao/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataVigenciaInicial",
            dataInicial: "2023-01-01",
            dataFinal: today,
            nivel: ["1", "2"],
            tipo: ["0", "2", "4"],
            status: ["0", null, "1", "3", "4", "5", "6", "7"],
          }),
        }
      );

      if (!response.ok) {
        console.log(`Erro na requisição: ${response.status}`);
        break;
      }

      const result = await response.json();
      nextPage = result.links.next;
      const producaoData = result.data;

      for (const producao of producaoData) {
        allResults.push(producao);
      }

      console.log(`Iterando... ${allResults.length}`);

      page++;

      await delay(2000);
    } while (nextPage !== null);

    console.log(`Quantidade total de produções: ${allResults.length}`);
    qtdItens = allResults.length

    return allResults;
  } catch (err) {
    console.log(`Erro: ${err}`);
  }
}

async function saveToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("tabelaProducao");

  const columns = [
    { header: "Proposta ID", key: "propostaId", width: 15 },
    { header: "Data Vigência Inicial", key: "dataVigenciaInicial", width: 20 },
    { header: "Data Vigência Final", key: "dataVigenciaFinal", width: 20 },
    { header: "Nível", key: "nivel", width: 10 },
    { header: "Nível Label", key: "nivelLabel", width: 20 },
    { header: "Tipo", key: "tipo", width: 15 },
    { header: "Tipo Label", key: "tipoLabel", width: 20 },
    { header: "Status Label", key: "statusLabel", width: 20 },
    { header: "Comissao", key: "comissao", width: 20 },
    { header: "Prêmio Líquido", key: "premioLiquido", width: 20 },
    { header: "Prêmio Total", key: "premioTotal", width: 20 },
    { header: "Parcelas", key: "parcelas", width: 20 },
    { header: "Nome do Corretor", key: "corretoresNome", width: 20 },
    { header: "Nome do Segurado", key: "seguradoNome", width: 20 },
    { header: "Segurado PF/PJ", key: "seguradoTipoPessoaLabel", width: 20 },
    {
      header: "Quantidade de Produções Segurado",
      key: "seguradoQtdeRegistrosProducao",
      width: 20,
    },
    { header: "Sexo do Segurado", key: "seguradoSexoLabel", width: 20 },
    { header: "Ramo", key: "ramoNome", width: 20 },
    { header: "Seguradora", key: "companhiaNome", width: 20 },
  ];

  worksheet.columns = columns;

  data.forEach((item) => {
    worksheet.addRow({
      propostaId: item.propostaId || "",
      dataVigenciaInicial: item.dataVigenciaInicial || "",
      dataVigenciaFinal: item.dataVigenciaFinal || "",
      nivel: item.nivel || "",
      nivelLabel: item.nivelLabel || "",
      tipo: item.tipo || "",
      tipoLabel: item.tipoLabel || "",
      statusLabel: item.statusLabel || "",
      comissao: item.comissao || "",
      premioLiquido: item.premioLiquido || "",
      premioTotal: item.premioTotal || "",
      parcelas: item.parcelas || "",
      corretoresNome: item.corretores[0].nome || "",
      seguradoNome: item.segurado.nome || "",
      seguradoTipoPessoaLabel: item.segurado.tipoPessoaLabel || "",
      seguradoQtdeRegistrosProducao: item.segurado.qtdeRegistrosProducao || "",
      seguradoSexoLabel: item.segurado.sexoLabel || "",
      ramoNome: item.ramo.nome || "",
      companhiaNome: item.companhia.nome || "",
    });
  });

  workbook.xlsx
    .writeFile("tabelaProducao.xlsx")
    .then(() => {
      console.log("Arquivo Excel gerado com sucesso!");
    })
    .catch((err) => {
      console.error("Erro ao gerar o arquivo Excel:", err);
    });
}

async function main() {
  const producaoData = await getProducao();
  await saveToExcel(producaoData);

  const email = process.env.MAIL_EMAIL;
  const password = process.env.MAIL_PASSWORD;
  var dateAndHour = `${String(date.getFullYear())}-${String(
    date.getMonth() + 1
  ).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")} às ${String(date.getHours())}:${String(date.getMinutes())}:${String(date.getSeconds())}`;  


  let transporter = nodemailer.createTransport({
    host: "smtp.hostinger.com",
    port: 465,
    auth: {
      user: email,
      pass: password
    }
  })

  transporter.sendMail({
    from: email,
    to: email,
    subject: 'A Tabela de Produções foi Atualizada!',
    html: `<h1 style="color: green">Produções Atualizadas com Sucesso!</h1>
    <p>Segue os dados:</p>
    <ul>
    <li>Quantidade de Produções: ${qtdItens}</li>
    <li>Atualizado dia ${dateAndHour}</li>
    </ul>
    `
  })
  .then(() => {
    console.log('E-mail PRODUÇÃO Enviado!')
  })
  .catch((err) => {
    console.log(`Erro ao enviar o e-mail ${err}`)
  })
}

main();
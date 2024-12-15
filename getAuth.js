import dotenv from "dotenv";
dotenv.config();

export let infoAPI = {
  url: "https://apirest.gruposgcor.com.br/api",
  email: process.env.TOKEN_EMAIL,
  password: process.env.TOKEN_PASSWORD,
};

export async function getAuth() {
  try {
    const response = await fetch(`${infoAPI.url}/login`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        email: infoAPI.email,
        senha: infoAPI.password,
      }),
    });

    if (response.ok) {
      const result = await response.json();
      return result;
    } else {
      console.log(`Erro na requisição: ${response.status}`);
    }
  } catch (err) {
    console.log(`Erro: ${err}`);
  }
}
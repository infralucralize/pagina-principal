const { ConfidentialClientApplication } = require("@azure/msal-node");
const fs = require("fs");
const path = require("path");

// ── Vem dos GitHub Actions Secrets, nunca hardcoded ──
const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;     // c0868f3b-764c-4c5b-a9fc-4af4b6eb0baf (lucralize-gestao-comercial)
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const CLIENTES = {
  siteId: "lucralize.sharepoint.com,1b12427a-fb1d-4615-bc54-9428543c1caa,9770018b-72f7-4925-860d-87996ece53f4",
  listId: "b0f38233-984e-4d0d-8dbd-3f5b13b63e91",
};
const USUARIOS = {
  siteId: "lucralize.sharepoint.com,8da49a9f-c97e-4d15-9706-a0fee47e966c,cb2492de-6af1-4063-bcf5-9eae681eafc8",
  listId: "a24fa86e-afda-4588-9ff7-410f1590bb64",
};

// Nome interno da coluna "Situação" — igual nas duas listas
const CAMPO_SITUACAO = "Situa_x00e7__x00e3_o";
const STATUS_VALIDO = "ATIVO";

const META_FIXED_VALUE = Number(process.env.META_FIXED_VALUE || 585);

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  },
});

async function getAppToken() {
  const result = await msalClient.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

// Conta só os itens com Situação = ATIVO. Precisa trazer os campos (expand=fields)
// porque agora filtramos por valor, não só contamos itens brutos.
async function countAtivos(token, { siteId, listId }) {
  let count = 0;
  let nextUrl =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items` +
    `?expand=fields(select=${CAMPO_SITUACAO})&$top=999`;

  while (nextUrl) {
    const resp = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) {
      const body = await resp.text();
      throw new Error(`Graph API respondeu ${resp.status} para siteId=${siteId}: ${body}`);
    }
    const json = await resp.json();
    for (const item of json.value) {
      if (item.fields && item.fields[CAMPO_SITUACAO] === STATUS_VALIDO) {
        count++;
      }
    }
    nextUrl = json["@odata.nextLink"] || null;
  }
  return count;
}

async function main() {
  const token = await getAppToken();

  const [totalClientesContabilidade, totalClientesTech] = await Promise.all([
    countAtivos(token, CLIENTES),
    countAtivos(token, USUARIOS),
  ]);

  const data = {
    totalClientesContabilidade,
    totalClientesTech,
    total: totalClientesContabilidade + totalClientesTech,
    meta: META_FIXED_VALUE,
    atualizadoEm: new Date().toISOString(),
  };

  const outPath = path.join(__dirname, "..", "data", "dados.json");
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify(data, null, 2));
  console.log("Salvo em", outPath, data);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

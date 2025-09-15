import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

import { ConfidentialClientApplication } from '@azure/msal-node';
import axios from "axios";

const clientId;
const tenantId;
const clientSecret;
const scopes = ["https://graph.microsoft.com/.default"];
const SITE_ID = "m365x49507934.sharepoint.com,7f979cef-fd8b-4bde-8138-03c9b73e2bde,c1a4f3e1-cd14-4d76-8412-01760787a802"
const LIST_ID = "40d19dcd-f408-4304-914d-a33737fed295"
const BASE_URL = "https://graph.microsoft.com/v1.0/sites"


const authClient = new ConfidentialClientApplication({
  auth: {
    clientId,
    clientSecret,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  }
});


export const authgraph = async () => {
  try {
    const authResult = await authClient.acquireTokenByClientCredential({ scopes });
    return authResult ? authResult.accessToken : null;
  } catch (err) {
    console.error('Authentication failed:', err);
    return null;
  }
};



async function fetchSharePointListData() {
  try {
    const token = await authgraph();
    const endpoint = `${BASE_URL}/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
    if (!token) {
      console.warn('No auth token retrieved');
      return null;
    }
    const response = await axios.get(endpoint, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    if (response.status === 200) return response.data.value
    return null

  } catch (error) {
    console.error('Error fetching data:', error);
    return null
  }

}

const server = new McpServer({
  name: "supportRequestMCP",
  description: "A server that provides Support Requests Data",
  version: "1.0.0",
  tools: [

    {
      name: "get-support-request-data",
      description: "Get all the support request details from sharepoint list",
      parameters: {},
    },

  ],
});

// // Get Chuck Norris joke tool
// const getChuckJoke = server.tool(
//   "get-chuck-joke",
//   "Get a random Chuck Norris joke",
//   async () => {
//     const response = await fetch("https://api.chucknorris.io/jokes/random");
//     const data = await response.json();
//     return {
//       content: [
//         {
//           type: "text",
//           text: data.value,
//         },
//       ],
//     };
//   }
// );

// // Get Chuck Norris joke categories tool
// const getChuckCategories = server.tool(
//   "get-chuck-categories",
//   "Get all available categories for Chuck Norris jokes",
//   async () => {
//     const response = await fetch("https://api.chucknorris.io/jokes/categories");
//     const data = await response.json();
//     return {
//       content: [
//         {
//           type: "text",
//           text: data.join(", "),
//         },
//       ],
//     };
//   }
// );

// // Get Dad joke tool
// const getDadJoke = server.tool(
//   "get-dad-joke",
//   "Get a random dad joke",
//   async () => {
//     const response = await fetch("https://icanhazdadjoke.com/", {
//       headers: {
//         Accept: "application/json",
//       },
//     });
//     const data = await response.json();
//     return {
//       content: [
//         {
//           type: "text",
//           text: data.joke,
//         },
//       ],
//     };
//   }
// );

// // Get Yo Mama joke tool
// const getYoMamaJoke = server.tool(
//   "get-yo-mama-joke",
//   "Get a random Yo Mama joke",
//   async () => {
//     const response = await fetch(
//       "https://www.yomama-jokes.com/api/v1/jokes/random"
//     );
//     const data = await response.json();
//     return {
//       content: [
//         {
//           type: "text",
//           text: data.joke,
//         },
//       ],
//     };
//   }
// );



server.tool(
  "get-support-request-data",
  "Get all the support request details from sharepoint list",

  async () => {
    const dataRes = await fetchSharePointListData();

    if (!dataRes) {
      return {
        content: [
          {
            type: "text",
            text: "Failed to retrieve expense data",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(dataRes),
        },
      ],
    };
  },
);

const app = express();

// to support multiple simultaneous connections we have a lookup object from
// sessionId to transport
const transports: { [sessionId: string]: SSEServerTransport } = {};

app.get("/sse", async (req: Request, res: Response) => {
  // Get the full URI from the request
  const host = req.get("host");

  const fullUri = `https://${host}/message`;
  const transport = new SSEServerTransport(fullUri, res);

  transports[transport.sessionId] = transport;
  res.on("close", () => {
    delete transports[transport.sessionId];
  });
  await server.connect(transport);
});

app.post("/message", async (req: Request, res: Response) => {
  const sessionId = req.query.sessionId as string;
  const transport = transports[sessionId];
  if (transport) {
    await transport.handlePostMessage(req, res);
  } else {
    res.status(400).send("No transport found for sessionId");
  }
});

app.get("/", (_req, res) => {
  res.send("The Support Request MCP server is running!");
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
});

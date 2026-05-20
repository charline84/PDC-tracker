import axios from "axios";

async function run() {
  const res = await axios.get("https://recherche-entreprises.api.gouv.fr/search?q=TotalEnergies&per_page=1");
  console.log(JSON.stringify(res.data.results[0], null, 2));
}
run();

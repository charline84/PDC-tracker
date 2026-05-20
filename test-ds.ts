import https from 'https';
https.get('https://www.data.gouv.fr/api/1/datasets/689c42fa521ccf80ce954f83/', (res) => {
  let data = '';
  res.on('data', chunk => data += chunk);
  res.on('end', () => console.log(JSON.parse(data).resources.map(r => r.url)));
});

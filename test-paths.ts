import https from 'https';

const check = (url) => {
  return new Promise(resolve => {
    https.get(url, (res) => {
      resolve({url, status: res.statusCode});
    }).on('error', () => resolve({url, status: 'error'}));
  })
}

async function run() {
  const ids = ['9db13a09-72a9-4871-b430-13872b4890b3'];
  const prefixes = [
    'https://www.data.gouv.fr/api/1/datasets/resources',
    'https://www.data.gouv.fr/api/2/datasets/resources',
    'https://www.data.gouv.fr/api/1/resources',
    'https://www.data.gouv.fr/api/2/resources',
    'https://data.gouv.fr/api/1/datasets/resources',
    'https://data.gouv.fr/api/1/resources',
  ];
  
  for (const prefix of prefixes) {
     const url = `${prefix}/${ids[0]}/data/json/`;
     console.log(await check(url));
  }
}
run();

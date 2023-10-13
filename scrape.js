const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const FILE_NAME = '1-10000.xlsx';

const scrape = async () => {
    let index = 590
    let end = 10000
    let cnt = 0

    while (index <= end) {
        await axios.get(`https://apps.calbar.ca.gov/attorney/Licensee/Detail/${index + 1}`).then((res) => {
            const $ = cheerio.load(res.data);
            if ($(`b:contains("#${index + 1}")`).text() !== '') {
                const workbook = XLSX.readFile(FILE_NAME);
                const worksheet = workbook.Sheets['Sheet1'];

                var nameAndNumber = $(`b:contains("#${index + 1}")`).text().trim();
                let arr = nameAndNumber.split('#');
                const name = arr[0].trim();
                const id = arr[1].trim();

                var addressRaw = $('p:contains("Address:")').text().trim();
                const address = addressRaw.split(":")[1].trim();

                var phoneAndFax = $('p:contains("Phone:")').text().trim();
                arr = phoneAndFax.split("|");
                const phone = arr[0].split(":")[1].trim();
                const fax = arr[1].split(":")[1].trim();

                var emailAndWeb = $('p:contains("Email:")').text().trim();
                arr = emailAndWeb.split("|");
                const email = arr[0].split(":")[1].trim();
                const website = arr[1].split(":")[1].trim();

                // A[*]: No, B[*]: Number, C[*]: Name, D[*]: Address, E[*]: Phone, F[*]: Fax, G[*]: Email, H[*]: Website
                let cell = XLSX.utils.decode_cell(`A${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`A${cnt + 2}`].v = cnt + 1;

                cell = XLSX.utils.decode_cell(`B${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`B${cnt + 2}`].v = name;

                cell = XLSX.utils.decode_cell(`C${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`C${cnt + 2}`].v = id;

                cell = XLSX.utils.decode_cell(`D${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`D${cnt + 2}`].v = address;

                cell = XLSX.utils.decode_cell(`E${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`E${cnt + 2}`].v = phone;

                cell = XLSX.utils.decode_cell(`F${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`F${cnt + 2}`].v = fax;

                cell = XLSX.utils.decode_cell(`G${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`G${cnt + 2}`].v = email;

                cell = XLSX.utils.decode_cell(`H${cnt + 2}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`H${cnt + 2}`].v = website;

                XLSX.writeFile(workbook, FILE_NAME);
                cnt++;
            } else {
                console.log(`${index + 1} is NO active`);
            }
        }).catch(err => {
            console.log(err, index + 1);
            return;
        })
        index++;
    }
}

scrape();
const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const FILE_NAME = '0-20000.xlsx';

const scrape = async () => {
    let index = 0
    let end = 20000
    let cnt = 0

    while (index < end) {
        await axios.get(`https://apps.calbar.ca.gov/attorney/Licensee/Detail/${index + 1}`).then((res) => {
            const $ = cheerio.load(res.data);
            if ($(`b:contains("#${index + 1}")`).text() !== '') {
                const workbook = XLSX.readFile(FILE_NAME);
                const worksheet = workbook.Sheets['Sheet1'];

                var nameAndNumber = $(`b:contains("#${index + 1}")`).text().trim();
                let arr = nameAndNumber.split('#');
                const name = arr[0].trim();
                const id = arr[1].trim();

                var addressRaw = $('p:contains("Address:")').text() !== '' ? $('p:contains("Address:")').text().trim() : '';
                const address = addressRaw !== '' ? addressRaw.split(":")[1].trim() : '';

                var phoneAndFax = $('p:contains("Phone:")').text() !== '' ? $('p:contains("Phone:")').text().trim() : '';
                arr = $('p:contains("Phone:")').text() !== '' ? phoneAndFax.split("|") : '';
                const phone = arr !== '' ? arr[0].split("Phone:")[1].trim() : '';
                const fax = arr !== '' ? arr[1] ? arr[1].split("Fax:")[1].trim() : '' : '';

                var emailAndWeb = $('p:contains("Email:")').text() !== '' ? $('p:contains("Email:")').text().trim() : '';
                arr = $('p:contains("Email:")').text() !== '' ? emailAndWeb.split("|") : '';
                let email = ''
                if (arr !== '') {
                    let id = ''
                    let index = res.data.indexOf(`{display:inline;}`)
                    if (res.data[index - 3] === 'e') id = res.data.substring(index - 3, index)
                    else id = res.data.substring(index - 2, index);
                    email = $(`span#${id}`).text()
                }
                const website = arr !== '' ? arr[1] ? arr[1].split("Website:")[1].trim() : '' : '';

                const ele = $('tbody tr:first-child')
                const status = ele.find('span').text() ? ele.find('span').text().trim() : '';

                const ele1 = $('tbody tr:last-child')
                const admissionDate = ele1.find('strong').text() ? ele1.find('strong').text().trim() : '';

                // A[*]: No, B[*]: Number, C[*]: Name, D[*]: Address, E[*]: Phone, F[*]: Fax, G[*]: Email, H[*]: Website
                let cell = XLSX.utils.decode_cell(`A${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`A${cnt + 1}`].v = cnt + 1;

                cell = XLSX.utils.decode_cell(`B${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`B${cnt + 1}`].v = name;

                cell = XLSX.utils.decode_cell(`C${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`C${cnt + 1}`].v = id;

                cell = XLSX.utils.decode_cell(`D${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`D${cnt + 1}`].v = address;

                cell = XLSX.utils.decode_cell(`E${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`E${cnt + 1}`].v = phone;

                cell = XLSX.utils.decode_cell(`F${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`F${cnt + 1}`].v = fax;

                cell = XLSX.utils.decode_cell(`G${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`G${cnt + 1}`].v = email;

                cell = XLSX.utils.decode_cell(`H${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`H${cnt + 1}`].v = website;

                cell = XLSX.utils.decode_cell(`I${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`I${cnt + 1}`].v = status;

                cell = XLSX.utils.decode_cell(`J${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`J${cnt + 1}`].v = admissionDate;

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
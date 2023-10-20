const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const FILE_NAME = 'ids.xlsx';
const RESULT_FILE_NAME = 'scraping.xlsx';

const getIds = async () => {
    let index = 0
    let end = 1775
    let cnt = 0

    while (index < end) {
        await axios.get(`https://www.mywsba.org/PersonifyEbusiness/Default.aspx?TabID=1536&ShowSearchResults=TRUE&EligibleToPractice=Y&Page=${index}`).then((res) => {
            const $ = cheerio.load(res.data);
            const workbook = XLSX.readFile(FILE_NAME);
            const worksheet = workbook.Sheets['Sheet1'];

            const trs = $('table.search-results tr.grid-row');
            for (let i = 0; i < trs.length; i++) {
                const tr = $(trs[i]);
                const value = tr.attr('onclick');
                const id = value.split("Usr_ID=")[1].substring(0, value.split("Usr_ID=")[1].length - 1)

                let cell = XLSX.utils.decode_cell(`A${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`A${cnt + 1}`].v = cnt + 1;

                cell = XLSX.utils.decode_cell(`B${cnt + 1}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`B${cnt + 1}`].v = id;

                cnt++;
            }

            XLSX.writeFile(workbook, FILE_NAME);
        }).catch(err => {
            console.log(err, index + 1, cnt);
            return;
        })
        index++;
    }
}

const scrape = async () => {
    const idWorkBook = XLSX.readFile(FILE_NAME);
    const idWorkSheet = idWorkBook.Sheets['Sheet1'];
    let index = 31573;
    let cnt = 1;

    while (idWorkSheet[`B${index}`]) {
        const id = idWorkSheet[`B${index}`].v;

        await axios.get(`https://www.mywsba.org/PersonifyEbusiness/LegalDirectory/LegalProfile.aspx?Usr_ID=${id}`).then((res) => {
            const $ = cheerio.load(res.data);
            const workbook = XLSX.readFile(RESULT_FILE_NAME);
            const worksheet = workbook.Sheets['Sheet1'];

            const sections = $('div.section')

            // Profile Data
            const profileSecetion = $(sections[0]);
            const name = profileSecetion.find('span').first().text().trim(); //B

            let licenseNumber = null; //C
            let licenseType = null; //D
            let llltArea = null; //E
            let eligibleToPractice = null; //F
            let licenseStatus = null; //G
            let wsbaAdmitDate = null; //H

            let address = null; //I
            let email = null; //J
            let phone = null; //K
            let fax = null; //L
            let website = null; //M
            let tdd = null; //N

            let firm = null; //O
            let office = null; //P
            let area = null; //Q
            let lang = null; //R

            let privatePractice = null; //S
            let has = null; //T
            let last_updated = null; //U

            let servedAs = null; //V
            let committees = null; //W

            let trs = profileSecetion.find('table tr')
            for (let i = 0; i < trs.length; i++) {
                const tr = $(trs[i])
                const header = tr.find('td').first().text().trim();
                const text = tr.find('td').last().text().trim();

                switch (header) {
                    case 'License Number:':
                        licenseNumber = text;
                        break;
                    case 'License Type:':
                        licenseType = text;
                        break;
                    case 'LLLT Practice Areas:':
                        llltArea = text;
                        break;
                    case 'Eligible To Practice:':
                        eligibleToPractice = text;
                        break;
                    case 'License Status:':
                        licenseStatus = text;
                        break;
                    case 'WSBA Admit Date:':
                        wsbaAdmitDate = text;
                        break;
                    default:
                        break;
                }
            }

            for (i = 1; i < sections.length; i++) {
                const section = $(sections[i]);
                const sectionHeader = section.find('span').first().text().trim();

                switch (sectionHeader) {
                    case 'Contact Information':
                        trs = section.find('table tr')
                        for (let i = 0; i < trs.length; i++) {
                            const tr = $(trs[i])
                            const header = tr.find('td').first().text().trim();
                            const text = tr.find('td').last().text().trim();

                            switch (header) {
                                case 'Public/Mailing Address:':
                                    address = text.split('\n').map(t => t.trim()).join(' ');
                                    break;
                                case 'Email:':
                                    email = text;
                                    break;
                                case 'Phone:':
                                    phone = text;
                                    break;
                                case 'Fax:':
                                    fax = text;
                                    break;
                                case 'Website:':
                                    website = text;
                                    break;
                                case 'TDD':
                                    tdd = text;
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;
                    case 'Practice Information Identified by Legal Professional':
                        trs = section.find('table tr')
                        for (let i = 0; i < trs.length; i++) {
                            const tr = $(trs[i])
                            const header = tr.find('td').first().text().trim();
                            const text = tr.find('td').last().text().trim();

                            switch (header) {
                                case 'Firm or Employer:':
                                    firm = text;
                                    break;
                                case 'Office Type and Size:':
                                    office = text;
                                    break;
                                case 'Practicse Areas:':
                                    area = text;
                                    break;
                                case 'Languages Other than English:':
                                    lang = text;
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;
                    case 'Professional Liability Insurance':
                        trs = section.find('table tr')
                        for (let i = 0; i < trs.length; i++) {
                            const tr = $(trs[i])
                            const header = tr.find('td').first().text().trim();
                            const text = tr.find('td').last().text().trim();

                            switch (header) {
                                case 'Private Practice:':
                                    privatePractice = text;
                                    break;
                                case 'Has Insurance?':
                                    has = text.split('-')[0].trim();
                                    break;
                                case 'Last Updated:':
                                    last_updated = text;
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;
                    case 'Judicial Service':
                        trs = section.find('table tr');
                        for (let i = 0; i < trs.length; i++) {
                            const tr = $(trs[i])
                            const header = tr.find('td').first().text().trim();
                            const text = tr.find('td').last().text().trim();

                            switch (header) {
                                case 'Has Ever Served as Judge:':
                                    servedAs = text;
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;
                    case 'Committees':
                        committees = section.find('p').first().next().text().trim();
                        break;
                    default:
                        break;
                }
            }

            let cell = XLSX.utils.decode_cell(`A${cnt}`);
            XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
            worksheet[`A${cnt}`].v = cnt;

            cell = XLSX.utils.decode_cell(`B${cnt}`);
            XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
            worksheet[`B${cnt}`].v = name;

            if (licenseNumber) {
                cell = XLSX.utils.decode_cell(`C${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`C${cnt}`].v = licenseNumber;
            }

            if (licenseType) {
                cell = XLSX.utils.decode_cell(`D${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`D${cnt}`].v = licenseType;
            }

            if (llltArea) {
                cell = XLSX.utils.decode_cell(`E${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`E${cnt}`].v = llltArea;
            }

            if (eligibleToPractice) {
                cell = XLSX.utils.decode_cell(`F${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`F${cnt}`].v = eligibleToPractice;
            }
            
            if (licenseStatus) {
                cell = XLSX.utils.decode_cell(`G${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`G${cnt}`].v = licenseStatus;
            }

            if (wsbaAdmitDate) {
                cell = XLSX.utils.decode_cell(`H${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`H${cnt}`].v = wsbaAdmitDate;
            }

            if (address) {
                cell = XLSX.utils.decode_cell(`I${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`I${cnt}`].v = address;
            }

            if (email) {
                cell = XLSX.utils.decode_cell(`J${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`J${cnt}`].v = email;
            }

            if (phone) {
                cell = XLSX.utils.decode_cell(`K${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`K${cnt}`].v = phone;
            }

            if (fax) {
                cell = XLSX.utils.decode_cell(`L${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`L${cnt}`].v = fax;
            }

            if (website) {
                cell = XLSX.utils.decode_cell(`M${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`M${cnt}`].v = website;
            }
            
            if (tdd) {
                cell = XLSX.utils.decode_cell(`N${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`N${cnt}`].v = tdd;
            }

            if (firm) {
                cell = XLSX.utils.decode_cell(`O${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`O${cnt}`].v = firm;
            }

            if (office) {
                cell = XLSX.utils.decode_cell(`P${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`P${cnt}`].v = office;
            }

            if (area) {
                cell = XLSX.utils.decode_cell(`Q${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`Q${cnt}`].v = area;
            }

            if (lang) {
                cell = XLSX.utils.decode_cell(`R${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`R${cnt}`].v = lang;
            }

            if (privatePractice) {
                cell = XLSX.utils.decode_cell(`S${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`S${cnt}`].v = privatePractice;
            }

            if (has) {
                cell = XLSX.utils.decode_cell(`T${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`T${cnt}`].v = has;
            }

            if (last_updated) {
                cell = XLSX.utils.decode_cell(`U${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`U${cnt}`].v = last_updated;
            }

            if (servedAs) {
                cell = XLSX.utils.decode_cell(`V${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`V${cnt}`].v = servedAs;
            }

            if (committees) {
                cell = XLSX.utils.decode_cell(`W${cnt}`);
                XLSX.utils.sheet_add_aoa(worksheet, [['']], { origin: cell });
                worksheet[`W${cnt}`].v = committees;
            }

            XLSX.writeFile(workbook, RESULT_FILE_NAME);
            console.log(index + " Scrapped");
        }).catch(err => {
            console.log(err, index);
            return;
        })
        index++;
        cnt++;
    }
}

scrape();
const axios = require('axios');
const fs = require('fs');
const FormData = require('form-data');
const xlsx = require('xlsx');

const API_KEY = "OcMPHc25bMlYCNfMc4mWV7FyZS8vViPqXT7z32HS";
const IMAGE_PATH = "./IMG_5786.JPG";
const OUTPUT_EXCEL_PATH = "./driver_duty_roster.xlsx";

async function validateApiKey(apiKey) {
    try {
        const response = await axios.get('https://validator.extracttable.com/', {
            headers: { 'x-api-key': apiKey }
        });
        console.log("API Key Usage Info:", response.data);
    } catch (error) {
        console.error("Error validating API key:", error.response?.data || error.message);
        throw error;
    }
}

async function triggerJob(apiKey, imagePath) {
    try {
        const formData = new FormData();
        formData.append('input', fs.createReadStream(imagePath));

        const response = await axios.post('https://trigger.extracttable.com/', formData, {
            headers: {
                ...formData.getHeaders(),
                'x-api-key': apiKey
            }
        });

        console.log("Job triggered successfully:", response.data);
        return response.data.JobId;
    } catch (error) {
        console.error("Error triggering job:", error.response?.data || error.message);
        throw error;
    }
}

async function getJobResult(apiKey, jobId) {
    try {
        while (true) {
            const response = await axios.get(`https://getresult.extracttable.com/?JobId=${jobId}`, {
                headers: { 'x-api-key': apiKey }
            });

            const jobStatus = response.data.JobStatus;
            console.log("Job Status:", response.data);

            if (jobStatus === "Success") {
                const tables = response.data.Tables;

                // Check if a downloadable file URL is provided
                if (response.data.FileUrl) {
                    const fileUrl = response.data.FileUrl;
                    console.log("Downloading file from:", fileUrl);

                    const fileResponse = await axios.get(fileUrl, { responseType: 'stream' });
                    const outputFilePath = './extracted_table_data.zip'; // Save as a ZIP file
                    const writer = fs.createWriteStream(outputFilePath);

                    fileResponse.data.pipe(writer);

                    await new Promise((resolve, reject) => {
                        writer.on('finish', resolve);
                        writer.on('error', reject);
                    });

                    console.log(`File downloaded successfully: ${outputFilePath}`);
                }

                // Retain the first three rows of data
                tables.forEach(table => {
                    if (table.TableJson) {
                        const tableJson = Object.values(table.TableJson);
                        table.TableJson = tableJson.slice(0, 3).concat(tableJson.slice(3)); // Keep first 3 rows + rest
                    }
                });

                return tables;
            } else if (jobStatus === "Failed") {
                throw new Error("Job failed to process the file.");
            } else if (jobStatus === "Incomplete") {
                console.warn("Job incomplete. Partial data may be available.");
                return response.data.Tables;
            }

            console.log("Waiting for job to complete...");
            await new Promise(resolve => setTimeout(resolve, 15000)); // Wait 15 seconds before retrying
        }
    } catch (error) {
        console.error("Error retrieving job result:", error.response?.data || error.message);
        throw error;
    }
}

function saveToJson(tables, outputPath) {
    try {
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth() + 1; // Current month (1-12)
        const currentYear = currentDate.getFullYear();

        const shiftMapping = {
            "A": "6:00 - 14:00",
            "A1": "10:00 - 18:00",
            "A2": "12:00 - 20:00",
            "A3": "08:00 - 16:00",
            "A4": "09:00 - 17:00",
            "B": "14:00 - 22:00",
            "B1": "16:00 - 00:00",
            "C": "22:00 - 6:00",
            "C1": "20:00 - 04:00",
            "C2": "18:00 - 02:00",
            "OFF": null
        };

        const jsonData = tables.map((table, index) => {
            const tableJson = table.TableJson;
            if (!tableJson || typeof tableJson !== 'object') {
                console.warn(`Table ${index + 1} has no valid data.`);
                return [];
            }

            return Object.values(tableJson)
                .filter((row, rowIndex) => rowIndex >= 3 && row[1] === "Bean") // Start from row 4 and filter rows where column B is "Bean"
                .map((row, rowIndex) => {
                    const name = row[0]; // Assume the name is in the first column
                    let dayCounter = 21; // Start from the 21st of the previous month
                    let month = currentMonth - 1;
                    let year = currentYear;

                    // Map each column (starting from E) to a JSON object
                    return Object.values(row).slice(4).map((cell) => {
                        let upperCell = typeof cell === 'string' ? cell.toUpperCase() : cell;

                        // Replace specific values
                        if (upperCell === "AL" || upperCell === "AI") upperCell = "A1";
                        if (upperCell === "BL" || upperCell === "BI") upperCell = "B1";
                        if (upperCell === "CL" || upperCell === "CI") upperCell = "C1";
                        if (upperCell === "OFFL" || upperCell === "OFFI") upperCell = "OFFL1";

                        const isOffWithNumber = upperCell && upperCell.startsWith("OFF") && !isNaN(upperCell.slice(3));

                        const result = {
                            name: name || "Unknown", // Add name field
                            date: `${dayCounter}/${month > 12 ? 1 : month}/${year}`,
                            shift: isOffWithNumber ? upperCell : (upperCell || "OFF"), // Retain OFF<số> or mark as "OFF"
                            time: isOffWithNumber ? null : (shiftMapping[upperCell] || null) // Default to "Nghỉ" for off days
                        };

                        dayCounter++;
                        if (dayCounter > 31) {
                            dayCounter = 1;
                            month++;
                            if (month > 12) {
                                month = 1;
                                year++;
                            }
                        }

                        return result;
                    });
                }).flat();
        }).flat();

        fs.writeFileSync(outputPath, JSON.stringify(jsonData, null, 2));
        console.log(`JSON file saved at: ${outputPath}`);
    } catch (error) {
        console.error("Error saving to JSON:", error.message);
        throw error;
    }
}

async function main() {
    try {
        await validateApiKey(API_KEY);

        // const jobId = await triggerJob(API_KEY, IMAGE_PATH);

        const tables = await getJobResult(API_KEY, "OcMPHc25bMlYCNfMc4mWV7FyZS8vViPqXT7z32HS_0knTaQbWBC_1755412582_I");
        // console.log("Tables retrieved:", JSON.stringify(tables));

        if (tables && tables.length > 0) {
            saveToJson(tables, "./driver_duty_roster.json");
        } else {
            console.log("No table data found in the image.");
        }
    } catch (error) {
        console.error("Error:", error.message);
    }
}

main();


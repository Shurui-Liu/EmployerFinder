// Global variables
let foundCompanies = [];
let existingCompanies = [];

// Initialize the add-in when Office.js is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('searchButton').onclick = searchCompanies;
        document.getElementById('addToExcelButton').onclick = addToExcel;
        document.getElementById('clearResultsButton').onclick = clearResults;
        document.getElementById('dismissErrorButton').onclick = hideError;
        
        // Initialize the add-in
        console.log('Company Search AI add-in loaded successfully');
    }
});

// Main search function
async function searchCompanies() {
    const searchCriteria = document.getElementById('searchCriteria').value.trim();
    const companyCount = parseInt(document.getElementById('companyCount').value);
    const apiKey = document.getElementById('apiKey').value.trim();

    // Validate inputs
    if (!searchCriteria) {
        showError('Please enter search criteria');
        return;
    }

    if (isNaN(companyCount) || companyCount < 1 || companyCount > 50) {
        showError('Please enter a valid number of companies (1-50)');
        return;
    }

    // Show loading state
    setLoadingState(true);
    showStatus('Reading existing companies from Excel...');

    try {
        // Read existing companies from Excel
        await readExistingCompanies();
        
        showStatus('Searching for companies using AI...');
        
        // Search for new companies using AI
        const newCompanies = await searchCompaniesWithAI(searchCriteria, companyCount, apiKey);
        
        // Filter out duplicates
        const uniqueCompanies = filterDuplicates(newCompanies);
        
        if (uniqueCompanies.length === 0) {
            showStatus('No new companies found that match your criteria');
            setLoadingState(false);
            return;
        }

        // Display results
        displayResults(uniqueCompanies);
        showStatus(`Found ${uniqueCompanies.length} new companies`);
        setLoadingState(false);

    } catch (error) {
        console.error('Error searching companies:', error);
        showError(`Error searching companies: ${error.message}`);
        setLoadingState(false);
    }
}

// Read existing companies from Excel first column
async function readExistingCompanies() {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = worksheet.getUsedRange();
                const firstColumn = usedRange.getColumn(0);
                const values = firstColumn.values;
                
                // Extract company names from the first column
                existingCompanies = values
                    .flat()
                    .filter(value => value && typeof value === 'string' && value.trim() !== '')
                    .map(value => value.trim().toLowerCase());
                
                console.log(`Found ${existingCompanies.length} existing companies`);
                resolve();
            } catch (error) {
                reject(error);
            }
        });
    });
}

// Search companies using AI API
async function searchCompaniesWithAI(criteria, count, apiKey) {
    const prompt = `Find ${count} real companies that match the following criteria: "${criteria}". 
    Return only the company names, one per line, without any additional text, numbers, or formatting. 
    Focus on well-known, legitimate companies that would be suitable for business research.`;

    try {
        // Use OpenAI API if key is provided, otherwise use a fallback
        if (apiKey) {
            return await callOpenAI(prompt, apiKey);
        } else {
            return await callFallbackAPI(prompt, count);
        }
    } catch (error) {
        throw new Error(`AI search failed: ${error.message}`);
    }
}

// Call OpenAI API
async function callOpenAI(prompt, apiKey) {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: 'gpt-3.5-turbo',
            messages: [
                {
                    role: 'system',
                    content: 'You are a helpful assistant that finds real companies based on criteria. Return only company names, one per line.'
                },
                {
                    role: 'user',
                    content: prompt
                }
            ],
            max_tokens: 500,
            temperature: 0.7
        })
    });

    if (!response.ok) {
        throw new Error(`OpenAI API error: ${response.status}`);
    }

    const data = await response.json();
    const companies = data.choices[0].message.content
        .split('\n')
        .map(line => line.trim())
        .filter(line => line && !line.startsWith('-') && !line.startsWith('•'))
        .slice(0, 20); // Limit to 20 companies

    return companies;
}

// Fallback API (simulated for demo purposes)
async function callFallbackAPI(prompt, count) {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Return sample companies based on common criteria
    const sampleCompanies = {
        'tech': ['Apple Inc.', 'Microsoft Corporation', 'Google LLC', 'Amazon.com Inc.', 'Meta Platforms Inc.', 'Netflix Inc.', 'Tesla Inc.', 'Salesforce Inc.', 'Adobe Inc.', 'Oracle Corporation'],
        'manufacturing': ['General Electric', 'Boeing Company', 'Ford Motor Company', 'General Motors', 'Caterpillar Inc.', '3M Company', 'Honeywell International', 'United Technologies', 'Lockheed Martin', 'Raytheon Technologies'],
        'healthcare': ['Johnson & Johnson', 'Pfizer Inc.', 'UnitedHealth Group', 'Merck & Co.', 'Abbott Laboratories', 'Medtronic plc', 'Amgen Inc.', 'Gilead Sciences', 'Bristol-Myers Squibb', 'Eli Lilly and Company'],
        'finance': ['JPMorgan Chase & Co.', 'Bank of America', 'Wells Fargo & Company', 'Citigroup Inc.', 'Goldman Sachs Group', 'Morgan Stanley', 'American Express', 'BlackRock Inc.', 'Charles Schwab', 'Visa Inc.'],
        'retail': ['Walmart Inc.', 'Target Corporation', 'Costco Wholesale', 'Home Depot Inc.', 'Lowe\'s Companies', 'Best Buy Co.', 'Starbucks Corporation', 'McDonald\'s Corporation', 'Nike Inc.', 'Coca-Cola Company']
    };

    // Determine category based on prompt
    let category = 'tech'; // default
    const promptLower = prompt.toLowerCase();
    
    if (promptLower.includes('manufacturing') || promptLower.includes('industrial')) {
        category = 'manufacturing';
    } else if (promptLower.includes('health') || promptLower.includes('medical') || promptLower.includes('pharma')) {
        category = 'healthcare';
    } else if (promptLower.includes('finance') || promptLower.includes('bank') || promptLower.includes('financial')) {
        category = 'finance';
    } else if (promptLower.includes('retail') || promptLower.includes('consumer') || promptLower.includes('shopping')) {
        category = 'retail';
    }

    return sampleCompanies[category].slice(0, count);
}

// Filter out duplicate companies
function filterDuplicates(newCompanies) {
    return newCompanies.filter(company => {
        const companyLower = company.toLowerCase();
        return !existingCompanies.some(existing => 
            existing.includes(companyLower) || companyLower.includes(existing)
        );
    });
}

// Display search results
function displayResults(companies) {
    foundCompanies = companies;
    const resultsList = document.getElementById('resultsList');
    
    resultsList.innerHTML = companies.map(company => 
        `<div class="result-item">${company}</div>`
    ).join('');
    
    document.getElementById('resultsSection').style.display = 'block';
}

// Add companies to Excel
async function addToExcel() {
    if (foundCompanies.length === 0) {
        showError('No companies to add');
        return;
    }

    setLoadingState(true);
    showStatus('Adding companies to Excel...');

    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = worksheet.getUsedRange();
            
            // Find the next empty row in column A
            const firstColumn = usedRange.getColumn(0);
            const values = firstColumn.values;
            const nextRow = values.length + 1;
            
            // Add companies to the next available rows
            for (let i = 0; i < foundCompanies.length; i++) {
                const cell = worksheet.getRange(`A${nextRow + i}`);
                cell.values = [[foundCompanies[i]]];
                
                // Apply green highlighting to new companies
                cell.format.fill.color = '#90EE90'; // Light green
                cell.format.font.bold = true;
            }
            
            await context.sync();
        });

        showStatus(`Successfully added ${foundCompanies.length} companies to Excel`);
        clearResults();
        setLoadingState(false);

    } catch (error) {
        console.error('Error adding to Excel:', error);
        showError(`Error adding to Excel: ${error.message}`);
        setLoadingState(false);
    }
}

// Clear results
function clearResults() {
    foundCompanies = [];
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('resultsList').innerHTML = '';
}

// Show error message
function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorSection').style.display = 'block';
    document.getElementById('statusSection').style.display = 'none';
}

// Hide error message
function hideError() {
    document.getElementById('errorSection').style.display = 'none';
}

// Show status message
function showStatus(message) {
    document.getElementById('statusMessage').textContent = message;
    document.getElementById('statusSection').style.display = 'block';
    document.getElementById('errorSection').style.display = 'none';
}

// Set loading state
function setLoadingState(isLoading) {
    const searchButton = document.getElementById('searchButton');
    const btnText = searchButton.querySelector('.btn-text');
    const btnLoading = searchButton.querySelector('.btn-loading');
    const progressBar = document.getElementById('progressBar');

    if (isLoading) {
        searchButton.disabled = true;
        btnText.style.display = 'none';
        btnLoading.style.display = 'flex';
        progressBar.style.display = 'block';
    } else {
        searchButton.disabled = false;
        btnText.style.display = 'block';
        btnLoading.style.display = 'none';
        progressBar.style.display = 'none';
    }
}

// Utility function to show success message
function showSuccess(message) {
    showStatus(message);
    setTimeout(() => {
        document.getElementById('statusSection').style.display = 'none';
    }, 3000);
} 
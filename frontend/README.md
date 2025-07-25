# Company Search AI - Excel Add-in

A lightweight Excel add-in that uses AI to search for companies based on specific criteria and automatically adds them to your spreadsheet with highlighting.

## Features

- 🔍 **AI-Powered Search**: Find companies using natural language criteria
- 📊 **Excel Integration**: Seamlessly add companies to your spreadsheet
- 🎨 **Visual Highlighting**: New companies are highlighted in green
- 🔄 **Duplicate Prevention**: Automatically filters out existing companies
- 📱 **Responsive UI**: Modern, clean interface that works on all devices
- ⚡ **Fast Performance**: Optimized for quick searches and updates

## How It Works

1. **Input**: Describe the type of companies you want to find
2. **Search**: AI analyzes your criteria and finds matching companies
3. **Filter**: Duplicates are automatically removed
4. **Add**: Companies are added to Excel with green highlighting
5. **Review**: New companies are clearly marked for easy identification

## Setup Instructions

### Prerequisites

- Microsoft Excel (desktop or online)
- A web server to host the add-in files
- Optional: OpenAI API key for enhanced AI capabilities

### Installation Steps

1. **Host the Files**
   - Upload all files to a web server (HTTPS required)
   - Update the URLs in `manifest.xml` to point to your server

2. **Install in Excel**
   - Open Excel
   - Go to Insert > Add-ins > My Add-ins
   - Choose "Upload My Add-in"
   - Select the `manifest.xml` file
   - Click "Upload"

3. **Configure API (Optional)**
   - Get an OpenAI API key from [OpenAI Platform](https://platform.openai.com/)
   - Enter the key in the add-in interface
   - If no key is provided, the add-in uses a fallback system

### File Structure

```
employer_search/
├── manifest.xml          # Add-in configuration
├── index.html           # Main UI
├── styles.css           # Styling
├── app.js              # Core functionality
├── commands.html       # Commands page
└── README.md           # This file
```

## Usage Guide

### Basic Usage

1. **Prepare Your Excel Sheet**
   - Ensure company names are in the first column (A)
   - The add-in will read existing companies to avoid duplicates

2. **Search for Companies**
   - Click the "Search Companies" button in the Home tab
   - Enter your search criteria (e.g., "tech startups in California")
   - Specify how many companies to find (1-50)
   - Click "Search Companies"

3. **Review and Add**
   - Review the found companies
   - Click "Add to Excel Sheet" to add them
   - New companies will be highlighted in green

### Search Criteria Examples

- "Tech companies with 1000+ employees"
- "Manufacturing companies in the Midwest"
- "Healthcare startups in Boston"
- "Financial services companies"
- "Sustainable energy companies in Europe"

### Advanced Features

- **Custom API Key**: Use your own OpenAI API key for better results
- **Duplicate Filtering**: Intelligent matching prevents duplicates
- **Visual Feedback**: Progress indicators and status messages
- **Error Handling**: Clear error messages for troubleshooting

## Technical Details

### AI Integration

The add-in supports two modes:

1. **OpenAI API** (with API key)
   - Uses GPT-3.5-turbo for intelligent company search
   - More accurate and contextual results
   - Requires API key and internet connection

2. **Fallback System** (no API key)
   - Pre-defined company lists by category
   - Works offline
   - Good for testing and basic use

### Excel Operations

- **Reading**: Extracts company names from column A
- **Writing**: Adds new companies to the next available rows
- **Formatting**: Applies green highlighting and bold text
- **Error Handling**: Graceful handling of Excel errors

### Security

- API keys are stored locally (not transmitted to external servers)
- HTTPS required for add-in hosting
- No data is stored or transmitted except for API calls

## Troubleshooting

### Common Issues

1. **Add-in Not Loading**
   - Ensure HTTPS hosting
   - Check manifest.xml URLs
   - Verify Office.js is accessible

2. **API Errors**
   - Check API key validity
   - Verify internet connection
   - Try fallback mode

3. **Excel Integration Issues**
   - Ensure Excel has write permissions
   - Check for protected worksheets
   - Verify Office.js compatibility

### Error Messages

- "Please enter search criteria" - Fill in the search field
- "Invalid number of companies" - Enter a number between 1-50
- "No new companies found" - Try different search criteria
- "Excel error" - Check worksheet permissions

## Development

### Customization

- **Styling**: Modify `styles.css` for custom appearance
- **AI Logic**: Update `app.js` for different AI providers
- **Excel Operations**: Extend Excel functionality in `app.js`

### Testing

1. Use the fallback API for testing
2. Test with various Excel file formats
3. Verify duplicate filtering works correctly
4. Check responsive design on different screen sizes

## Support

For issues and questions:
- Check the troubleshooting section
- Review browser console for errors
- Ensure all files are properly hosted
- Verify Excel add-in permissions

## License

This project is provided as-is for educational and business use.

---

**Note**: This add-in requires a web server with HTTPS to function properly in Excel. For production use, consider hosting on a reliable platform like Azure, AWS, or GitHub Pages. 
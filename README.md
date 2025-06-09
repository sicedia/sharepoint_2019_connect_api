# Nintex Connect API - SharePoint Connector

A Python application for connecting to SharePoint Server 2019 using REST API with NTLM authentication to retrieve list items and export them to CSV format.

## Features

- **NTLM Authentication**: Secure connection to SharePoint Server 2019
- **REST API Integration**: Uses SharePoint REST API for data retrieval
- **Pagination Support**: Automatically handles large datasets with pagination
- **Data Export**: Exports retrieved data to CSV format
- **Flexible Retrieval**: Support for both limited and unlimited item retrieval
- **Clean Data Processing**: Removes technical metadata from SharePoint items
- **Error Handling**: Comprehensive error handling and logging

## Requirements

- Python 3.8+
- SharePoint Server 2019 or compatible version
- Valid SharePoint credentials with list access permissions

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd nintex-connect-api
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root with your SharePoint configuration:
```env
SP_SITE_URL=https://your-sharepoint-site.com
SP_USERNAME=your-username
SP_PASSWORD=your-password
```

## Dependencies

```
pandas>=1.5.0
requests>=2.28.0
python-dotenv>=0.19.0
requests-ntlm>=1.1.0
```

## Usage

### Basic Usage

Run the application to retrieve all items from the configured SharePoint list:

```bash
python main.py
```

This will:
1. Connect to SharePoint using the credentials from `.env`
2. Retrieve all items from the list "RA4-1 Solicitud para Viajes"
3. Display results in the console
4. Save data to `all_items.csv`

### Programmatic Usage

```python
from main import SharePointConnector, SharePointConfig

# Create configuration
config = SharePointConfig(
    site_url="https://your-sharepoint-site.com",
    username="your-username",
    password="your-password",
    list_title="Your List Name"
)

# Initialize connector
connector = SharePointConnector(config)

# Get all items (with automatic pagination)
all_items = connector.get_all_items()

# Get limited items
limited_items = connector.get_items_with_limit(100)

# Save to CSV
all_items.to_csv("output.csv", index=False)
```

## Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `SP_SITE_URL` | SharePoint site URL | `https://sgi.cedia.org.ec` |
| `SP_USERNAME` | SharePoint username | `nintexinstall` |
| `SP_PASSWORD` | SharePoint password | (empty) |

### Application Constants

You can modify these constants in `main.py`:

- `LIST_TITLE`: The SharePoint list name to query
- `DEFAULT_TIMEOUT`: HTTP request timeout in seconds
- `REQUEST_HEADERS`: HTTP headers for SharePoint API requests

## Architecture

### Classes

#### `SharePointConfig`
Data class that holds SharePoint connection configuration:
- `site_url`: SharePoint site URL
- `username`: Authentication username
- `password`: Authentication password
- `list_title`: Target SharePoint list name

#### `SharePointConnector`
Main connector class with methods:
- `get_all_items()`: Retrieves all items with automatic pagination
- `get_items_with_limit(limit)`: Retrieves up to specified number of items
- Internal methods for URL building, session management, and data processing

### Key Features

- **Automatic Pagination**: Handles SharePoint's pagination automatically
- **Data Cleaning**: Removes SharePoint metadata fields (starting with `_`)
- **Error Handling**: Comprehensive HTTP error handling with logging
- **Session Management**: Reuses HTTP session for efficiency
- **Type Safety**: Full type hints for better code maintainability

## Output

The application generates:
1. **Console Output**: Summary of retrieved records and preview
2. **CSV File**: Complete dataset exported to `all_items.csv`

Example console output:
```
2025-06-09 10:30:15 | INFO     | root | Requesting URL: https://...
Retrieved 150 records.
   ID  Title           Created              Modified
0   1  Travel Request  2025-01-15T09:00:00  2025-01-15T09:00:00
1   2  Travel Request  2025-01-16T10:30:00  2025-01-16T10:30:00
...
2025-06-09 10:30:16 | INFO     | root | Data saved to all_items.csv
```

## Error Handling

The application handles common scenarios:
- **Network errors**: HTTP timeouts and connection issues
- **Authentication errors**: Invalid credentials
- **API errors**: SharePoint REST API errors
- **Data validation**: Invalid limit parameters
- **Empty results**: No items found in the list

## Logging

Logging is configured to show:
- Timestamp
- Log level
- Logger name
- Message

Log levels used:
- `INFO`: Normal operations (URL requests, data saved)
- `WARNING`: Non-critical issues (empty results)
- `ERROR`: Critical errors (HTTP failures)

## Compatibility

- **SharePoint Versions**: SharePoint Server 2019, SharePoint 2016, SharePoint 2013
- **Authentication**: NTLM authentication
- **API**: SharePoint REST API with OData support
- **Python**: Python 3.8+

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Verify credentials in `.env` file
   - Ensure user has access to the SharePoint list
   - Check if NTLM authentication is enabled

2. **List Not Found**
   - Verify the list title is correct (case-sensitive)
   - Ensure the list exists in the specified site
   - Check user permissions to access the list

3. **Connection Timeout**
   - Check network connectivity to SharePoint server
   - Verify the `SP_SITE_URL` is correct
   - Consider increasing `DEFAULT_TIMEOUT` value

4. **Empty Results**
   - Verify the list contains items
   - Check if user has read permissions
   - Review SharePoint list permissions

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review the logs for error details
3. Open an issue in the repository with:
   - Error messages
   - Environment details
   - Steps to reproduce
# Company Research Automation System

A Django-based application that automates company research by scraping data from Google Sheets, processing it with OpenAI, and storing results in both PostgreSQL database and formatted Google Sheets.

## 🚀 Features

- **Dual Storage**: Saves data to PostgreSQL database AND Google Sheets
- **AI-Powered Research**: Uses OpenAI GPT-3.5-turbo for company information extraction
- **Admin Panel**: Django admin interface for data management
- **Formatted Output**: Color-coded Google Sheets with structured sections
- **REST API**: Simple endpoints for triggering scraping and health checks
- **Docker Ready**: Containerized for easy deployment

## 📋 Prerequisites

- Python 3.11+
- PostgreSQL database (Neon recommended)
- Google Sheets API access
- OpenAI API key
- Docker (optional)

## 🛠️ Installation

### 1. Clone and Setup

```bash
git clone <repository>
cd company-research
cp .env.example .env
```

### 2. Configure Environment Variables

Edit `.env` file with your credentials:

```bash
# Database (Neon PostgreSQL)
DATABASE_URL=postgresql://username:password@host:port/database?sslmode=require

# APIs
OPEN_AI_API_KEY=sk-proj-your-openai-key
API_KEY=your-google-sheets-api-key
SPREADSHEET_ID=your-google-spreadsheet-id

# Google Service Account
PROJECT_ID=your-google-project-id
PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nyour-key\n-----END PRIVATE KEY-----"
CLIENT_EMAIL=your-service-account@project.iam.gserviceaccount.com
CLIENT_ID=your-client-id
```

### 3. Run with Docker (Recommended)

```bash
docker-compose up --build
```

### 4. Run Locally

```bash
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

## 🔧 Configuration Setup

### Google Sheets API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Google Sheets API
4. Create Service Account credentials
5. Download JSON key file
6. Share your Google Sheet with the service account email
7. Copy credentials to `.env` file

### OpenAI API Setup

1. Go to [OpenAI Platform](https://platform.openai.com/)
2. Create API key
3. Add to `.env` as `OPEN_AI_API_KEY`

### Database Setup (Neon)

1. Create account at [Neon](https://neon.tech/)
2. Create new database
3. Copy connection string to `.env` as `DATABASE_URL`

## 📊 Data Flow

```
Google Sheets (Company List) 
    ↓
Django Application
    ↓
OpenAI GPT-3.5 (Research)
    ↓
Dual Storage:
├── PostgreSQL Database (Structured data)
└── Google Sheets (Formatted reports)
```

## 🎯 Usage

### API Endpoints

- **GET /** - Health check and welcome message
- **POST /scrape/** - Trigger company research process
- **GET /health/** - Application health status
- **GET /admin/** - Django admin panel (login required)

### Admin Panel

Access at `http://localhost:8000/admin/`

**Default credentials** (if using setup script):
- Username: `admin`
- Password: `admin123`

**Features:**
- View/edit company data
- Manage executives and offices
- View research history
- Bulk operations

### Scraping Process

1. Reads company list from Google Sheets ("会社リスト" sheet)
2. For each company:
   - Extracts corporate number, name, address
   - Sends to OpenAI for detailed research
   - Saves to PostgreSQL database
   - Creates formatted Google Sheet with color-coded sections
   - Tracks changes in history

## 📁 Project Structure

```
company-research/
├── manage.py                 # Django management
├── requirements.txt          # Dependencies
├── .env                     # Environment variables
├── Dockerfile               # Container setup
├── docker-compose.yml       # Docker orchestration
├── company_research/        # Django project
│   ├── settings.py         # Configuration
│   ├── urls.py             # Main routing
│   └── wsgi.py             # WSGI application
└── research/               # Main application
    ├── models.py           # Database models
    ├── admin.py            # Admin interface
    ├── views.py            # API endpoints
    ├── scraper.py          # Core scraping logic
    └── migrations/         # Database migrations
```

## 🗄️ Database Schema

### Company Model
- Basic info (name, corporate number, representative)
- Financial data (revenue, employees, capital)
- Business details (industry, operations)
- Contact information

### Executive Model
- Position, name, reading (furigana)
- Linked to company

### Office Model
- Office locations and contact details
- Business content per location
- Linked to company

### Research History
- Tracks all data changes
- Corporate number, field, old/new values
- Timestamp

## 🎨 Google Sheets Output

Creates individual sheets per company with color-coded sections:

- **Blue**: Basic corporate information
- **Green**: Financial information  
- **Orange**: Business content
- **Pink**: Executive roster
- **Yellow**: Office locations
- **Gray**: Related URLs

## 🚀 Deployment

### Docker Deployment

```bash
# Build and run
docker-compose up -d --build

# View logs
docker-compose logs -f

# Stop
docker-compose down
```



### Environment Variables for Production

```bash
DEBUG=False
SECRET_KEY=your-production-secret-key
DATABASE_URL=your-production-database-url
```

## 🔍 Troubleshooting

### Common Issues

1. **Database Connection Error**
   - Check DATABASE_URL format
   - Verify network connectivity to Neon

2. **Google Sheets API Error**
   - Verify service account has access to spreadsheet
   - Check API key and credentials

3. **OpenAI API Error**
   - Verify API key is valid
   - Check API usage limits

4. **Migration Issues**
   ```bash
   python manage.py makemigrations
   python manage.py migrate
   ```

## 📝 API Response Examples

### Successful Scraping
```json
{
  "status": "success",
  "message": "Scraping completed successfully",
  "result": {
    "processed": 12,
    "message": "Successfully processed 12 companies"
  }
}
```

### Health Check
```json
{
  "status": "healthy"
}
```

## 🤝 Contributing

1. Fork the repository
2. Create feature branch
3. Make changes
4. Test thoroughly
5. Submit pull request

## 📄 License

This project is licensed under the MIT License.

## 🆘 Support

For issues and questions:
1. Check troubleshooting section
2. Review logs: `docker-compose logs`
3. Create GitHub issue with details
# django-scraper

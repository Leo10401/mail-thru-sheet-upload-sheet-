# Mail Thru Sheet

A powerful web application that allows you to upload Excel/CSV files, manage contacts, and send personalized emails to your contacts. Built with React, Node.js, and modern web technologies.

## Features

- 📤 Upload Excel/CSV files with contact information
- 👥 Manage and organize contacts from spreadsheets
- 🔍 Search through contacts by name or email
- ✉️ Send personalized emails to selected contacts
- 🎨 Beautiful, responsive UI with dark mode support
- 📱 Mobile-friendly interface

## Tech Stack

### Frontend
- React + Vite
- TailwindCSS for styling
- React Email Editor for email templates
- Lucide React for icons
- React Hot Toast for notifications

### Backend
- Node.js with Express
- Multer for file uploads
- Nodemailer for email functionality
- XLSX for spreadsheet processing
- MongoDB for data storage

## Prerequisites

- Node.js (v20 or higher)
- npm or yarn
- MongoDB (if using database storage)
- SMTP server credentials

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/mail-thru-sheet.git
cd mail-thru-sheet
```

2. Install dependencies for both frontend and backend:
```bash
# Install backend dependencies
cd backend
npm install

# Install frontend dependencies
cd ../frontend
npm install
```

3. Configure environment variables:
   - Create a `.env` file in the backend directory with the following variables:
   ```
   SMTP_HOST=your_smtp_host
   SMTP_PORT=your_smtp_port
   SMTP_USER=your_smtp_username
   SMTP_PASSWORD=your_smtp_password
   SMTP_SECURE=true_or_false
   ```

## Running the Application

### Development Mode

1. Start the backend server:
```bash
cd backend
npm run dev
```

2. Start the frontend development server:
```bash
cd frontend
npm run dev
```

### Production Mode

The application can be run using Docker Compose:

```bash
docker-compose up --build
```

This will start both the frontend and backend services.

## Usage

1. Upload your Excel/CSV file containing contact information
2. Select the appropriate sheet from your spreadsheet
3. Browse and search through your contacts
4. Create or select an email template
5. Send personalized emails to selected contacts

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

If you encounter any issues or have questions, please open an issue in the GitHub repository.

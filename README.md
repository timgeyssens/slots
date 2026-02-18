# School Feedback Slot Booking System

A simple slot management system for student feedback sessions, powered by Google Sheets.

## Features

- **Teacher Management**: Add teachers with their availability hours
- **Configurable Slots**: Set slot duration, breaks after X slots, and break duration
- **Sequential Booking**: Students can only book the first available slot (prevents gaps)
- **Google Sheets Backend**: All data stored in a Google Sheet you control

## Setup Instructions

### Option A: Setup Wizard (Recommended)

The app includes a built-in wizard that guides you through the setup:

1. **Run the frontend** (use a local server for best results):
   ```bash
   python3 -m http.server 8000
   ```
   Then open http://localhost:8000

2. **Click "Setup Wizard"** on the homepage

3. **Create Google Cloud credentials** (one-time, free):
   - Go to [Google Cloud Console](https://console.cloud.google.com/apis/credentials)
   - Create a project and OAuth client ID (Web application)
   - Add `http://localhost:8000` as authorized JavaScript origin
   - Enable the [Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com)

4. **Follow the wizard steps**:
   - Enter your Client ID
   - Sign in with Google
   - Create the sheet (automatic!)
   - Copy the Apps Script code and deploy it

### Option B: Manual Setup

#### 1. Create the Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it something like "Feedback Slots"

#### 2. Add the Apps Script

1. In your Google Sheet, go to **Extensions → Apps Script**
2. Delete any existing code in the editor
3. Copy the entire contents of `Code.gs` and paste it
4. Click **Save** (disk icon) and name the project "Slot Booking"

#### 3. Initialize the Sheets

1. In the Apps Script editor, select `initializeSheets` from the function dropdown
2. Click **Run**
3. Grant the necessary permissions when prompted
4. This creates the required sheets with headers

#### 4. Deploy as Web App

1. Click **Deploy → New deployment**
2. Click the gear icon and select **Web app**
3. Set the following:
   - **Description**: "Slot Booking API"
   - **Execute as**: Me
   - **Who has access**: Anyone
4. Click **Deploy**
5. Click **Authorize access** and grant permissions
6. **Copy the Web app URL** (looks like `https://script.google.com/macros/s/xxx/exec`)

#### 5. Run the Frontend

**Local server (recommended)**
```bash
# Using Python
python3 -m http.server 8000

# Using Node.js (if you have npx)
npx serve .
```
Then open http://localhost:8000

**Or host online** - Upload `index.html` to GitHub Pages, Netlify, etc.

#### 6. Connect to Google Sheets

1. Open the webpage
2. Click "Manual Setup"
3. Paste your Google Apps Script URL
4. Click **Save URL**

## Usage

### Admin (⚙️ Admin tab)

1. **Configure Slots**: Set duration (e.g., 15 min), slots before break (e.g., 4), and break duration (e.g., 15 min)
2. **Add Teachers**: Enter name and availability hours
3. **Manage Bookings**: View and cancel bookings

### Students (📅 Book a Slot tab)

1. View all teachers and their available slots
2. The **first available slot** for each teacher has a "Book Now" button
3. Enter name and email to confirm booking

## How It Works

- **Sequential booking enforced**: Students can only book the next available slot for each teacher
- **Automatic breaks**: After every X slots (configurable), a break is automatically inserted
- **Real-time data**: All changes sync immediately with Google Sheets
- **No server needed**: Frontend talks directly to Google Apps Script

## Customization

### Change Colors
Edit the CSS in `index.html` - look for:
- `.btn-primary` - main action buttons (green)
- `.teacher-card-header` - teacher card headers

### Add More Fields
1. Add columns to the Bookings sheet
2. Update `bookSlot()` and `getBookings()` in `Code.gs`
3. Update the booking modal in `index.html`

## Troubleshooting

**"Failed to load data"**
- Check that your Script URL is correct
- Make sure the Apps Script is deployed as a web app
- Verify permissions are set to "Anyone"

**CORS errors**
- Google Apps Script handles CORS automatically
- If issues persist, try a different browser or clear cache

**Slots not showing**
- Run `initializeSheets()` in Apps Script to create sheet structure
- Add at least one teacher in the Admin tab

## Data Structure

The Google Sheet has three tabs:

| Sheet | Columns |
|-------|---------|
| Teachers | Name, Available From, Available Until |
| Config | Setting Name, Value |
| Bookings | Teacher, Slot Time, Student Name, Student Email, Timestamp |

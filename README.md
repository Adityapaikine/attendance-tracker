# Attendance & Leave Management System with P2P Messenger

A desktop-based **Attendance, Leave Management, and Internal Messaging System** built with Python and Tkinter.

The application combines employee attendance tracking, leave management, user profiles, announcements, and a peer-to-peer internal messaging system into a single desktop application.

---

## 📌 Project Overview

This project is designed for organizations that need a lightweight desktop application for managing:

* Employee accounts
* Employee profiles
* Attendance records
* Leave applications
* Leave approvals
* Announcements
* Internal user-to-user messaging
* User roles and permissions
* Profile photos
* Dark/Light mode
* Excel-based data storage

The project consists primarily of two Python components:

```text
attendance_and_leave_tracker_v2.py
p2p_messenger.py
```

The two components are integrated so users can access messaging functionality alongside the attendance and leave management system.

---

## ✨ Main Features

### 👤 User Management

The system supports multiple users with different roles.

User information is maintained in the Excel `Users` worksheet.

Typical user information includes:

| Field           | Description                   |
| --------------- | ----------------------------- |
| `username`      | Unique login username         |
| `full_name`     | Employee's full name          |
| `department`    | Employee department           |
| `password_hash` | Securely stored password hash |
| `role`          | User permission level         |

Users can log in using their assigned credentials.

---

### 🔐 Role-Based Access

The application supports different levels of access.

Typical roles include:

* **Master/Super User**
* **Administrator**
* **Employee/User**

Depending on the user's role, different operations may be available.

Administrative users can perform management operations such as:

* Managing employee records
* Reviewing attendance
* Managing leave requests
* Editing employee information
* Managing announcements
* Accessing administrative functionality

Regular users can generally:

* View their own attendance
* Apply for leave
* View their leave information
* View announcements
* View/edit permitted profile information
* Send and receive internal messages

---

## 🕒 Attendance Management

The attendance module allows the organization to maintain employee attendance records.

Typical attendance information includes:

* Employee username
* Employee name
* Date
* Check-in time
* Check-out time
* Attendance status
* Additional attendance information

Users can view their own attendance records.

Administrators can manage and review attendance information for employees.

---

## 🏖️ Leave Management

The leave management module allows employees to submit leave requests.

Employees can:

1. Open the leave section.
2. Select the appropriate leave type.
3. Specify the required dates.
4. Enter a reason.
5. Submit the leave request.
6. Monitor the status of the request.

Administrators can review and process leave applications.

Possible leave statuses include:

```text
Pending
Approved
Rejected
```

---

## 📢 Announcements

The dashboard contains an announcement section for communicating important information to users.

Announcements can be used for:

* Company notices
* Holiday announcements
* Attendance reminders
* Policy updates
* General employee communication

The announcement functionality is designed to make important information available to all users.

---

## 💬 P2P Messenger

The project includes an internal peer-to-peer messaging system.

The messenger allows one user to send a message directly to another user.

Example:

```text
User 1
   |
   | Message
   v
User 2
```

The messaging system is designed to support:

* User discovery
* User-to-user messaging
* Message history
* Persistent messages
* Temporary network disconnection
* Logout/login scenarios
* Automatic refreshing

### Message Persistence

Messages should not depend solely on the sender being logged in.

For example:

```text
1. User 1 logs in.
2. User 1 sends a message to User 2.
3. User 1 logs out.
4. User 2 logs in.
5. User 2 can retrieve the previously sent message.
```

This requires messages to be persisted rather than remaining only in the sender's local application state.

---

## 🔄 Automatic Refresh

The application uses automatic refresh mechanisms to keep information synchronized.

Refresh functionality is important for:

* New messages
* User availability
* Attendance information
* Leave status
* Announcements
* Other shared application data

A manual/global refresh mechanism may also be provided where applicable.

---

## 🌐 User Discovery

The P2P messenger is intended to allow active users to discover one another automatically.

Conceptually:

```text
             ┌──────────────┐
             │   User 1     │
             └──────┬───────┘
                    │
                    │
       ┌────────────┼────────────┐
       │            │            │
       ▼            ▼            ▼
 ┌──────────┐ ┌──────────┐ ┌──────────┐
 │  User 2  │ │  User 3  │ │  User 4  │
 └──────────┘ └──────────┘ └──────────┘
```

The application should handle users becoming available or unavailable without requiring the entire application to restart.

---

## 👨‍💼 Employee Profiles

Employees have individual profiles.

Profile information may include:

* Full name
* Username
* Department
* Role
* Other employee information
* Profile photo

### Profile Editing Rules

Regular users should not be able to modify protected employee information.

Users can be allowed to modify their profile photo.

Administrative/Master users retain the ability to modify protected employee information.

This provides a separation between:

```text
Employee-controlled information
             +
Administrator-controlled information
```

---

## 🖼️ Profile Photos

Users can add a profile photo to their account.

The application can display the profile image in areas such as:

* User profile
* Dashboard
* Messenger
* Employee information

The profile-photo functionality is intended to improve user identification and the overall UI experience.

---

## 🌙 Dark Mode

The application supports a dark-mode interface.

The theme system controls major UI elements such as:

* Backgrounds
* Frames
* Buttons
* Text
* Entry fields
* Tables
* Dashboard elements
* Messenger interface

The objective is to ensure that switching between themes does not leave unwanted white or mismatched UI areas.

---

## 📊 Excel Data Storage

The project uses an Excel workbook for structured application data.

A typical workbook may contain worksheets such as:

```text
leaves_data.xlsx
│
├── Users
├── Attendance
├── Leaves
├── Announcements
└── Messages
```

The exact worksheet structure depends on the current application version.

### Users Worksheet

Example:

| username | full_name    | department | password_hash   | role |
| -------- | ------------ | ---------- | --------------- | ---- |
| user1    | Employee One | HR         | hashed_password | User |
| user2    | Employee Two | IT         | hashed_password | User |

### Attendance Worksheet

Stores employee attendance information.

### Leaves Worksheet

Stores leave applications and their statuses.

### Announcements Worksheet

Stores announcements displayed to users.

### Messages Worksheet

Stores persistent internal messages where supported by the current implementation.

---

## 📁 Recommended Project Structure

A recommended project structure is:

```text
Attendance_Leave_Management/
│
├── attendance_and_leave_tracker_v2.py
├── p2p_messenger.py
├── leaves_data.xlsx
├── README.md
│
├── data/
│   └── application_data.xlsx
│
├── profiles/
│   └── user profile images
│
├── logs/
│   └── application logs
│
└── requirements.txt
```

The actual directory structure can differ depending on deployment requirements.

---

## 🛠️ Technology Stack

### Programming Language

**Python 3**

### GUI

**Tkinter**

### Data Storage

**Microsoft Excel / `.xlsx`**

### Excel Library

**openpyxl**

### Networking

The P2P messenger uses Python networking functionality for user discovery and message communication, depending on the current implementation.

### Security

Passwords should be stored as hashes rather than plaintext passwords.

---

## 📦 Installation

### 1. Install Python

Install Python 3 on the target computer.

Verify the installation:

```bash
python --version
```

or:

```bash
py --version
```

---

### 2. Install Required Packages

Install the required Python packages:

```bash
pip install openpyxl
```

If the current implementation uses additional third-party packages, install them as specified by the project's `requirements.txt`.

---

### 3. Download/Copy the Project

Place the project files in the desired directory.

Example:

```text
C:\AttendanceLeaveManagement\
```

Ensure the Python files and Excel data file are accessible to the application.

---

### 4. Prepare the Excel Workbook

Create or place the required Excel workbook in the location expected by the application.

For example:

```text
leaves_data.xlsx
```

Make sure the required worksheets exist and their column names match what the Python application expects.

---

## ▶️ Running the Application

Run the main application:

```bash
python attendance_and_leave_tracker_v2.py
```

If the messenger is designed to run as a separate process/application:

```bash
python p2p_messenger.py
```

If the messenger is integrated into the main application, launch only the main application.

---

## 🔑 Login

Users log in using their registered username and password.

A typical authentication process is:

```text
Start Application
       ↓
Login Screen
       ↓
Validate Username
       ↓
Validate Password Hash
       ↓
Load User Profile
       ↓
Determine User Role
       ↓
Load Dashboard
```

---

## 🔒 Security Considerations

The following security practices should be maintained:

### Passwords

Never store plaintext passwords.

Use password hashing:

```text
Password
   ↓
Hash Function
   ↓
Password Hash
   ↓
Excel/User Database
```

### User Permissions

Do not rely only on hidden UI buttons for security.

Permission checks should also be performed when an operation is executed.

For example:

```python
if current_user["role"] == "Master":
    # allow administrative operation
```

### Messenger Security

If the messenger operates across a network, authentication and message validation should be considered.

Untrusted messages should not be executed as Python code or interpreted as executable commands.

---

## 🔄 Message Delivery Architecture

A robust messaging workflow should follow this general architecture:

```text
Sender
  │
  ▼
Create Message
  │
  ▼
Persist Message
  │
  ▼
Identify Recipient
  │
  ├───────────────┐
  │               │
Recipient Online  Recipient Offline
  │               │
  ▼               ▼
Deliver Now       Keep Stored
  │               │
  └───────┬───────┘
          ▼
     Recipient Login
          │
          ▼
    Retrieve Messages
          │
          ▼
     Display Message
```

The important principle is:

> **Message persistence should be independent of the sender's login session.**

---

## 🧪 Testing

Before deployment, test the following scenarios.

### Authentication

* [ ] Valid username/password
* [ ] Invalid password
* [ ] Invalid username
* [ ] Different user roles
* [ ] Logout
* [ ] Login again

### Attendance

* [ ] User can view own attendance
* [ ] Administrator can view employee attendance
* [ ] Attendance records persist after restart

### Leave

* [ ] User can submit leave
* [ ] Leave appears after submission
* [ ] Administrator can approve leave
* [ ] Administrator can reject leave
* [ ] User sees updated status

### Profiles

* [ ] User can view profile
* [ ] User can update profile photo
* [ ] User cannot modify protected employee information
* [ ] Master/Administrator can modify protected information

### Announcements

* [ ] Administrator can create/update announcements
* [ ] Users can see announcements
* [ ] Announcements persist after restart

### Messenger

* [ ] User 1 can discover User 2
* [ ] User 2 can discover User 1
* [ ] User 1 can send a message to User 2
* [ ] User 2 receives the message while online
* [ ] User 1 logs out
* [ ] User 2 logs in afterward
* [ ] User 2 can still see the previously sent message
* [ ] Messages survive application restart
* [ ] Temporary network disconnection does not permanently lose messages
* [ ] Multiple users can communicate independently

### UI

* [ ] Light mode works
* [ ] Dark mode works
* [ ] No unwanted white areas remain in dark mode
* [ ] Windows do not freeze during normal operation
* [ ] Refresh functionality works correctly

---

## ⚠️ Common Issues

### Users Cannot See Previous Messages

Possible causes include:

* Messages are stored only in memory.
* Message history is not persisted.
* Recipient history is not loaded during login.
* Sender and recipient identifiers do not match.
* The messenger does not synchronize after reconnecting.
* The Excel message worksheet is not being updated.
* The application is reading from a different data file than the sender.

The recommended design is to persist messages and load the recipient's pending/history messages when the recipient logs in.

---

### Application Freezes

Tkinter's main UI thread should not perform long-running network or file operations.

Operations that may block should be handled using appropriate background processing and then update the UI safely.

Examples include:

* Network discovery
* Network communication
* Large Excel operations
* Frequent synchronization
* File operations

---

### Dark Mode Leaves White Areas

This generally occurs when a widget was created using a hard-coded background such as:

```python
bg="white"
```

instead of using the active theme.

Prefer centralized theme values:

```python
bg=COLORS["background"]
```

or:

```python
bg=self.colors["surface"]
```

All major widgets should use the application's theme system.

---

## 🚀 Recommended Future Improvements

Potential improvements include:

### Database

Move from Excel to a proper database such as:

* SQLite
* PostgreSQL
* MySQL

A database provides better concurrency and reliability than Excel for a multi-user application.

### Messaging

Consider implementing:

* Message IDs
* Delivery status
* Read/unread status
* Timestamps
* Retry queues
* Offline message queues
* Message acknowledgements
* Duplicate-message protection

Example:

```text
Message ID
Sender
Recipient
Timestamp
Content
Status
```

### Authentication

Consider:

* Strong password policies
* Session management
* Account lockout
* Password reset
* Role-based authorization
* Secure credential storage

### Networking

For larger deployments, consider using a central server instead of pure peer-to-peer communication:

```text
             ┌───────────────┐
             │ Message Server│
             └───────┬───────┘
                     │
       ┌─────────────┼─────────────┐
       │             │             │
       ▼             ▼             ▼
    User 1         User 2        User 3
```

This makes offline message delivery and synchronization considerably easier.

---

## 📄 Data Backup

Because the application may use Excel files for important employee data, regular backups are strongly recommended.

At minimum, back up:

```text
leaves_data.xlsx
```

before performing major application changes.

A backup naming scheme could be:

```text
leaves_data_backup_2026-08-22.xlsx
```

---

## 🏗️ Deployment

For deployment to computers that do not have Python installed, the application can be packaged as a Windows executable using a packaging tool such as PyInstaller.

A typical command is:

```bash
pyinstaller --onefile --windowed attendance_and_leave_tracker_v2.py
```

If the application requires external files such as Excel workbooks, profile images, icons, or configuration files, those resources must also be included or placed in the expected application directory.

---

## 📌 Development Guidelines

When modifying the project:

1. Back up the current Python files.
2. Back up the Excel workbook.
3. Make one major change at a time.
4. Test login/logout.
5. Test application restart.
6. Test multiple users.
7. Test network disconnection.
8. Check the Tkinter console for exceptions.
9. Verify that Excel data is actually being persisted.
10. Test both Light and Dark modes.

---

## 📝 Project Status

The project is an actively developed desktop application combining:

```text
Attendance
    +
Leave Management
    +
Employee Profiles
    +
Announcements
    +
Role-Based Access
    +
P2P Messaging
    +
Automatic Refresh
    +
Theme Management
```

The primary development focus is reliability, persistent data, multi-user synchronization, and a consistent desktop user experience.

---

## 📜 License

Add the project's applicable license here.

Example:

```text
Copyright © 2026

All rights reserved unless otherwise specified by the project owner.
```

---

## 👨‍💻 Project Components

| Component                            | Purpose                                                                     |
| ------------------------------------ | --------------------------------------------------------------------------- |
| `attendance_and_leave_tracker_v2.py` | Main attendance, leave, profile, dashboard, and user-management application |
| `p2p_messenger.py`                   | Internal peer-to-peer messaging functionality                               |
| `leaves_data.xlsx`                   | Excel-based application data storage                                        |
| `README.md`                          | Project documentation                                                       |

---

## ✅ Quick Start

```text
1. Install Python 3
        ↓
2. Install required packages
        ↓
3. Place application files together
        ↓
4. Prepare the Excel workbook
        ↓
5. Start attendance_and_leave_tracker_v2.py
        ↓
6. Log in
        ↓
7. Use Attendance / Leave / Profile / Announcements
        ↓
8. Open Messenger
        ↓
9. Discover users
        ↓
10. Send and receive messages
```

---

## 📞 Support / Maintenance

When reporting an issue, include:

* Python version
* Windows version
* Application version
* Exact error message
* Steps that reproduce the problem
* Whether the issue affects one user or all users
* Whether the issue occurs after logout/login
* Whether the issue occurs after restarting the application
* Relevant console traceback, if available

This information makes debugging significantly easier.

WhatsApp Business Suite
A desktop application designed to automate and streamline WhatsApp messaging for business purposes, including bulk sending, custom templates, and courier notifications.

Table of Contents
Features

Setup and Installation

Running the Application

Packaging the Application for Distribution

Creating a New Release on GitHub

Conclusion

Features
Multi-User Login: Secure login system with 'admin' and 'user' roles.

Admin Panel: Manage users and application settings from a dedicated admin interface.

Bulk Sender: Send templated or custom messages to a list of contacts from manual entry or an Excel file.

Courier Notifications: Automatically parse a formatted Excel sheet to send detailed courier dispatch notifications to customers.

Custom Field Sender: Upload any Excel file, map columns to variables (like {name}, {order_id}), and send personalized messages.

Template Manager: Create, save, and reuse message templates for any of the sending tools.

Image Support: Attach images to your bulk messages for richer content.

Forced Updates: Ensures users are always on the latest version by enforcing mandatory updates via a non-dismissible screen.

Setup and Installation
This project uses a virtual environment to manage dependencies, which keeps the project self-contained and ensures the final packaged application is as small as possible.

Create a Virtual Environment:
Open your terminal in the project's root directory and run:

python -m venv venv

Activate the Virtual Environment:

On Windows (PowerShell):

.\venv\Scripts\Activate.ps1

On Windows (Command Prompt):

.\venv\Scripts\activate.bat

On macOS/Linux:

source venv/bin/activate

You will see (venv) at the beginning of your terminal prompt.

Install Required Packages:
With the virtual environment active, install all necessary libraries using the requirements.txt file:

pip install -r requirements.txt

Running the Application
After installing the requirements, you can run the application directly from the source code:

python app.py

Packaging the Application for Distribution
To create a single .exe file that you can send to users, we use PyInstaller.

Make sure your virtual environment is active.

Run the following command in your terminal:

pyinstaller --name "WhatsApp Business Suite" --onefile --windowed --add-data "templates;templates" --add-data "chromedriver.exe;." app.py

The final, distributable .exe file will be located inside the dist/ folder. It's recommended to zip this file before sharing it.

Creating a New Release on GitHub
To make your packaged app available for download, follow these steps:

Navigate to your private repository on GitHub.

Click on the "Releases" link on the right-hand side.

Click on "Create a new release" or "Draft a new release".

Create a New Tag

The tag should match the version specified in your app.py. For example, if your LATEST_APP_VERSION is 2.0.0, name your tag v2.0.0.

Title Your Release

Give the release a title, such as "Version 2.0.0 - Major Update".

Write a Description

Provide a short description of the changes made in this release.

Attach Binaries

In the "Attach binaries" box, drag and drop your zipped .exe file (e.g., WhatsApp-Business-Suite.zip).

Publish Release

Click on "Publish release" to make it available.

Conclusion
Follow this workflow to maintain a smooth development process and ensure that your releases are well-documented and easily accessible for users.
# ClassRM Finder for Semmelweis University

ClassRM Finder is a Discord bot designed to help students and staff at Semmelweis University find classroom information and update calendar and Excel files accordingly.

## Table of Contents
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Commands](#commands)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)


## Features
- Retrieve classroom information and update Excel and calendar files
- User-friendly Discord bot interface
- Simple and fast classroom search functionality

## Prerequisites
- Python 3.7 or higher
- Discord account
- Discord bot token

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/chococobyeol/ClassRM_finder_4_semmelweis_univ.git
    cd ClassRM_finder_4_semmelweis_univ
    ```

2. Create a virtual environment:
    ```sh
    python3 -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

4. Set up your Discord bot:
    - Create a new bot on the [Discord Developer Portal](https://discord.com/developers/applications).
    - Copy the bot token and create a `.env` file in the project's root directory with the following content:
        ```
        DISCORD_TOKEN=your-bot-token
        ```

## Usage

1. Prepare your calendar (.ics) and Excel (.xlsx) files.

2. Run the bot:
    ```sh
    python main.py
    ```

3. Invite the bot to your Discord server using the OAuth2 URL generated on the Discord Developer Portal.

## Commands

### `!검색 <강의실 코드>`
- Example: `!검색 AOKIK`
- Searches for the specified classroom name and provides information.

### `!엑셀검색`
- Example: `!엑셀검색`
- Prompts the user to upload an Excel file (.xlsx) and searches for classroom information contained in the file, providing an updated file.

### `!캘린더검색`
- Example: `!캘린더검색`
- Prompts the user to upload a calendar file (.ics) and searches for classroom information contained in the file, providing an updated file.

### `!시간표변환`
- Example: `!시간표변환`
- Prompts the user to upload an Excel file (.xlsx) and converts it into a specified format, providing the converted file.

## Contributing

Contributions are always welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For any inquiries, please contact:
- Email: chococobyeol@gmail.com

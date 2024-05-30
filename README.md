# GhostGenerator

A tool to generate PowerPoint presentations using Gemini

## How it works

This tool uses a proxy so basically it uses the openai api for communicating in the library but has a custom url to send requests which then translates it to the gemini api

## Installation Instructions

Follow the steps below to set up and run the GhostGenerator;

### 1. Clone the Repository

Begin by cloning the repository to your local machine:

```bash
git clone https://github.com/The-UnknownHacker/GhostGenerator.git
cd GhostGenerator
```

## 2 -  Install the required packages using `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 3. Set up the .env File

The repository contains a `.env.example` file. You need to save this as `.env` and then update it with you Gemini API key:

First, make a copy and rename:

```bash
cp .env.example .env
```

Then, open the `.env` file in a text editor of your choice and replace `GEMINI_API_KEY` with your actual Gemini API key.

### 4. Run the Application

Once everything is set up, run the application using:

```bash
python main.py
```

---


<h1 style="font-size:3em; color:#1D6F42;">Excel-MCP</h1>

<p style="font-size:1.2em;">
  <strong>Your AI-powered Excel Automation Assistant</strong><br>
  <em>Let AI handle your Excel tasks—just use commands!</em>
</p>

<p>
  <a href="https://claude.ai/download" style="background:#0078D7; color:white; padding:8px 16px; border-radius:5px; text-decoration:none; font-weight:bold;">Download Claude Desktop</a><br>
  <span style="color:#777;">Connect to this MCP server, or build your own custom client for seamless interaction.</span>
</p>

## What is Excel-MCP?

Excel-MCP is a **tool that allows AI to interact directly with your Excel files**.  
If you are a beginner or intermediate in Excel and find formulas, sorting, or complex tasks overwhelming, let AI take care of the work.  

It allows you to:  
- Read and modify Excel sheets using natural language commands.  
- Automate repetitive operations such as filling formulas, filtering, or summarizing data.  
- Generate new reports, charts, or visualizations instantly.  
- Combine multiple Excel sheets or datasets automatically.  

### Example Commands for an LLM

Here are a few examples of what you can ask Claude Desktop to do with Excel-MCP:  

- `"Show all rows in Sheet1 that have missing values"`  
- `"Add a new column 'Total' which is the sum of columns A and B"`  
- `"Filter Sheet2 to show only rows where 'Status' is 'Completed'"`  
- `"Generate a bar chart of 'Sales' by 'Region' from Sheet3"`  
- `"Merge Sheet1 and Sheet2 on column 'CustomerID'"`  
- `"Highlight all values in column 'Revenue' greater than 10000"`  

---

## Features

- **Seamless Integration:** Works natively with Microsoft Excel.  
- **AI Automation:** Automate tedious and repetitive tasks with simple commands.  
- **Advanced Data Processing:** Powerful tools for data analysis and visualization.  
- **Security First:** Robust & secure architecture for peace of mind.  
- **Customizable:** Easily extend features to fit your workflow.  

---

## 🛠️ Installation

**Step 1: Clone the Repository**
```sh
git clone https://github.com/Mayank-Golchha/Excel-MCP.git
cd Excel-MCP

**Step 2: Set Up the Virtual Environment**

<details>
  <summary><strong>Windows</strong></summary>

  ```sh
  python -m venv .venv
  .venv\Scripts\activate
  pip install -r requirements.txt
  ```
</details>

<details>
  <summary><strong>Mac/Linux</strong></summary>

  ```sh
  python -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt
  ```
</details>

**Step 3: Run the Server**
```sh
uv run --with pandas excel_mcp_server.py
```

---

## Get Started

- Simply connect using the Claude Desktop app or a custom client.
- Start giving natural language commands for Excel tasks.
- Let AI take care of the magic while you focus on what matters!

---

<p align="center">
  <img src="https://img.shields.io/badge/Made%20with%20%E2%9D%A4%20by%20Mayank-Golchha-blue" alt="Made by Mayank-Golchha">
</p>

<p align="center">
  <a href="https://github.com/Mayank-Golchha/Excel-MCP/issues">🐞 Report Issues</a> | 
  <a href="https://github.com/Mayank-Golchha/Excel-MCP/pulls">💡 Contribute</a>
</p>

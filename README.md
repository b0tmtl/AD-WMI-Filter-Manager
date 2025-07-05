# AD WMI Filter Manager

<div align="center">

![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=for-the-badge&logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-Custom%20Non--Commercial-yellow.svg?style=for-the-badge)
![Version](https://img.shields.io/badge/Version-3.1-blue.svg?style=for-the-badge)

<h3>A modern PowerShell GUI tool for managing and testing Active Directory WMI Filters</h3>

[Features](#features) • [Installation](#installation) • [Usage](#usage) • [Requirements](#requirements) • [License](#license)

</div>

---

## 📋 Overview

AD WMI Filter Manager is a professional-grade PowerShell tool that provides a modern graphical interface for managing Windows Management Instrumentation (WMI) filters in Active Directory environments. Built with WPF and featuring a sleek dark theme, this tool simplifies the complex task of viewing, testing, and managing WMI filters across your domain.

## ✨ Features

### Core Functionality
- 🔍 **Browse WMI Filters** - View all WMI filters in your domain with instant search
- 🔗 **GPO Linking** - See which Group Policy Objects are linked to each filter
- 🧪 **Live Testing** - Test WMI filters against specific computers in real-time
- 🔐 **Credential Support** - Test filters using alternate credentials when needed
- ⚡ **Performance Optimized** - Fast loading and responsive interface

### User Experience
- 🎨 **Modern Dark Theme** - Easy on the eyes with a professional appearance
- 🔎 **Instant Search** - Filter results as you type
- 📊 **Statistics Dashboard** - See filter counts and GPO links at a glance
- 💫 **Smooth Animations** - Polished interactions and transitions
- 🖱️ **Intuitive Interface** - No training required

## 🖼️ Screenshots

<div align="center">
![Main Interface](https://github.com/user-attachments/assets/b5a01867-d14e-493a-9aac-050da0d7dac2)
</div>

## 🚀 Installation

### Option 1: Direct Download
1. Download `WmiGUI.ps1` from the [Releases](https://github.com/YOUR-USERNAME/AD-WMI-Filter-Manager/releases) page
2. Right-click the file and select "Properties"
3. Check "Unblock" if present and click "Apply"
4. Run the script with PowerShell

### Option 2: Git Clone
```powershell
git clone https://github.com/b0tmtl/AD-WMI-Filter-Manager.git
cd AD-WMI-Filter-Manager
.\WmiGUI.ps1

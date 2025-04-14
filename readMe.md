# 📊 Spring Boot Excel Report Generator

This app filters daily support tickets and generates a detailed Excel report with subtasks, ticket aging, and custom formatting.

## 🚀 Features
- REST endpoints for Excel generation
- Reads input + grid Excel files
- Outputs a detailed report with colored sections (CSR, DevOps, etc.)
- Calculates working-day-based ticket age

## 📡 API
- `GET /filter-issues` → Filter and process `Service Request` issues

## 📂 Structure
- `Controller.java` — API endpoints
- `ExcelServices.java` — Filtering and writing logic
- `application.properties` — (Optional) configurable file paths

## 📥 Input Files
- `input.xlsx` — Ticket dump
- `Daily Report updated Grid.xlsx` — Project metadata

## 📦 Output File
- `/home/decimal/Documents/dd-mm-yyyy.xlsx`

## ✍️ Author
Created by [Vishwajeet Singh](https://github.com/yourusername)

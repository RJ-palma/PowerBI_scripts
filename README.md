# PowerShell Scripts Repository

Welcome to the PowerShell Scripts Repository! This collection of scripts is designed for various tasks including cloning, deleting, downloading, publishing, and checking the status of files. Below is a list of the available scripts with brief descriptions.

## Table of Contents
- [Cloning](#cloning)
- [Deleting](#deleting)
- [Downloading](#downloading)
- [Publishing](#publishing)
- [Status Checker](#status-checker)

## Cloning
### `workspace-cloningV2.ps1`
This script is used to clone a specified workspace/s. It allows for configuration of individual reports for the cloning process.

## Deleting
### `bulk-delete-files.ps1`
This script enables bulk deletion of Power BI reports. It can delete multiple files based on specified patterns or criteria.

## Downloading
### `bulk-add-filesV3.ps1`
This script facilitates the bulk downloading of Power BI reports. It can handle multiple download links and organize the downloaded Power BI reports accordingly.

## Publishing
### `bulk-publishingV2.ps1`
This script is utilized for bulk publishing of Power BI reports. It offers various options to customize how Power BI reports are published to different workspaces.

## Status Checker
### `fail-status-checker.ps1`
This script checks the status of various Power BI reports and their associated datasets. It helps in identifying any errors and provides insights into the availablitity of the Power BI reports.

## Usage
- Make sure you have the necessary permissions to execute these scripts.
- Install the following module in Powershell:
--- "Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser" ---
- Open PowerShell and run the script using the following command:
  ```powershell
  .\<folder_name>\<script-name>.ps1

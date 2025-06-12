# Pull Microsoft Store

This GitHub Action downloads application packages (`.msix`) for a given Microsoft Store Product ID directly from Microsoft's content delivery network.

It automates the complex process of communicating with the Microsoft Store's backend services to fetch download links for UWP and modern applications.

## How It Works

The action mimics the process a Windows device uses to download a store application. The high-level workflow is:

1.  **Fetch Product Metadata**: Resolves the provided `product-id` against the public Microsoft Store API to get package details, including its unique `WuCategoryId` and `PackageFamilyName`.
2.  **Authenticate**: Obtains a temporary authorization cookie from Microsoft's update servers. This is required for all subsequent requests.
3.  **Sync Updates**: Performs a `SyncUpdates` operation against the Windows Update service endpoint. It sends a detailed SOAP request identifying itself as a Windows client seeking updates for the specific application category.
4.  **Identify Files**: Parses the `SyncUpdates` response to identify all relevant file names and their corresponding `UpdateID` and `RevisionNumber`.
5.  **Fetch Download URLs**: For each file, it makes a `GetExtendedUpdateInfo2` request to retrieve the direct, CDN-hosted download URL.
6.  **Download and Save**: Downloads each file from its URL and saves it to the specified `output-path`.

## Usage

To use this action, add the following step to your workflow file (e.g., `.github/workflows/main.yml`).

### Example Workflow

This example downloads the packages for **Windows Terminal** and lists the contents of the output directory.

```yaml
name: Download Microsoft Store App Package
on:
  workflow_dispatch:

jobs:
  download-app:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout Code
        uses: actions/checkout@v4

      - name: Pull App Packages for Windows Terminal
        uses: sht2017/Pull-Microsoft-Store@main
        id: store_download
        with:
          product-id: '9NBLGGH4LS1F' # Product ID for Windows Terminal
          output-path: './app-packages' #optional

      - name: Verify Download
        if: steps.store_download.outcome == 'success'
        run: |
          echo "Successfully downloaded packages:"
          ls -R ./app-packages

      - name: Handle Failure
        if: steps.store_download.outcome != 'success'
        run: echo "Download failed."
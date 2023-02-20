## Requirements

This package requires the `graphviz` tool.

You can install Graphviz on MacOS via homebrew:

```bash
brew install graphviz
```

Or, if you are using Homestead:

```bash
sudo apt-get install graphviz
```

To install Graphviz on Windows, download it from the [official website](https://graphviz.gitlab.io/_pages/Download/Download_windows.html).

## Installation

You can install the package via composer:

```bash
composer require vinhnt/databasedocs:dev-main --dev
```
You need to add the following to `config\app.php`:

```php
\\ Register Service Providers
Vinhnt\Databasedocs\Providers\DatabaseDocsServiceProvider::class,
```
<!-- ## Usage

Run command 
```bash
php artisan generate:databasedocs
```
Excel file will be stored in `storage\app`

## Reference

https://oceanic-cut-66f.notion.site/Docs-Database-630f485e846a4e119382981064948067 -->
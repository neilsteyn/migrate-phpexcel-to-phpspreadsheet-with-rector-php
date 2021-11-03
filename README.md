# migrate-phpexcel-to-phpspreadsheet-with-rector-php
Migrating PhpExcel to PhpSpreadsheet with Rector PHP

ref: https://getrector.org/blog/2020/04/16/how-to-migrate-from-phpexcel-to-phpspreadsheet-with-rector-in-30-minutes

## 1. Install rector

composer install

## 2. Create a src directory and add all your PhpExcel code into the src directory

## 3.1 Dry run Rector to see what would be changed

php vendor/bin/rector process src --dry-run

## 3.2 Make the changes happen

php vendor/bin/rector process src

## 4. Change \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->xls, 'Excel5'); to \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->xls, 'Xls');


## 5. Load PhpSpreadsheet with composer autoload

require(str_replace('\\','/',dirname(__DIR__)).'/vendor/autoload.php'); or require('vendor/autoload.php');


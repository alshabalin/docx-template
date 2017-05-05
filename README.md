# Docx-Template

Allows convert any docx document to template, using variable, like `{this}`.

## Install the package

```bash
composer require alshabalin/docx-template
```

## How to use

```php
<?php

require 'vendor/autoload.php';

$data = [
  'key' => 'value',
  'name' => 'John',
  'lastname' => 'Doe',
  'city' => 'London',
];

$template = new DocxTemplate();

$template->open('document.docx')
    ->setData($data)
    ->save('result.docx');
```

You may want to remove any missing variables from your template by passing `true` as the second param to `setData` method:

```php
$template->setData($data, true);
```



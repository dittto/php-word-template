# PHP Word Template

## What is it?

A replacement for the template.php file for PHP Word (https://phpword.codeplex.com/).
This is based on code by http://jeroen.is/phpword-templates-with-repeating-rows/

The aim of this code is to allow much greater flexibility over tables. This code allows you create the rows of a table, remove rows and columns, and simple variables using the tag format `${tag_name}`.

There are some restrictions on tags (e.g. ${test}) in templates, and these revolve around how a docx is built. To write a tag, type it in one go, without deleting, without underscores or dashes, and without capital letters, or turn off spell-check in Word.

If a tag isn't being recognised then check the docx's source code and search for it. Chances are that it's being split into multiple xml tags in the docx file.

Most of the methods in this class are to be used as a toolbox.

## Example template output

An example template can be found in the file TestTemplate. The way to instantiate and save the file is:

```php
    $word = new Word_Word('save_name', '/dir/to/save/in');
    $template = new Word_TestTemplate($word, '/dir/to/template_file.docx');
    $template->createFile();
```

If you want to output the results you can do:

```php
    $word = new Word_Word('save_name', '/dir/to/save/in');
    $template = new Word_TestTemplate($word, '/dir/to/template_file.docx');
    $template->createFile();

    $headers = $word->getHeaders();
    foreach ($headers as $header) {
        header($header);
    }
    readfile('/dir/to/save/in/save_name.docx');
```

If you look in the TestTemplate file, you'll find the first method being called is `$this->repairTemplate()`. This is because .docx is a particularly poor format to use as it'll often split up text with tags for auto-correct flags another others.

This method will try to repair the .docx by removing unimportant tags. It's a good idea to run a new template through `repairTemplate()` once and then save the fixed version. You can then use that fixed version every time without the overhead of fixing it every time.

## Create a table

You can create a table from tags placed in the .docx template. Assuming you have the following setup in your template:

```
   First name      | Last name
   ${first_name}   | ${last_name}
```

You can use the following code in createFile() to create the table for as many rows as needed. The createRows only needs to be passed one tag that's in a table row:

```php
    $data = array();
    $data[] = array('first_name' => 'First', 'last_name' => 'Name');
    $data[] = array('first_name' => 'Another', 'last_name' => 'Name');
    $this->createRows('first_name', $data);
```

## Remove a row or column

Sometimes you'll need optional rows or columns, such as in the following tables:

```
    First name    | Last name    | Cost    | Has Paid
    ${first_name} | ${last_name} | ${cost} | ${has_paid}

    Subtotal: | ${subtotal}
    Discount: | ${discount}
    Total:    | ${total}
```

If we assume that in some invoices, such as the one above, you don't always need the discount row or the has_paid column, you can remove them by doing the following:

```php
    $this->removeRow('discount');
    $this->removeRow('has_paid');
```

## Duplicate a paragraph

Just altering the text in a paragraph is easy enough, just do:

```php
    $this->setTag('terms_header', 'Terms and conditions');
```

If you need more though, like duplicating bullet-pointed paragraphs, then you can do the following:

```php
    $this->cloneParagraph('terms', array('Term 1', 'Term 2'));
```
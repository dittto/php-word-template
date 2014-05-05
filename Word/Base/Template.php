<?php
/**
 * Class Word_Base_Template
 *
 * Creates a docx based on a template. This is based on
 * http://jeroen.is/phpword-templates-with-repeating-rows/
 *
 * There are some big restrictions on tags (e.g. ${test}) in templates, and
 * these revolve around how a docx is built. To write a tag, type it in one
 * go, without deleting, without underscores or dashes, and without capital
 * letters, or turn off spell-check in Word.
 *
 * If a tag isn't being recognised then check the docs source code and search
 * for it. Chances are that it's being split into multiple xml tags in the
 * docx file.
 */
abstract class Word_Base_Template {
    /**
     * A store for the Word_Word object
     * @var Word_Base_Base
     */
    private $word;

    /**
     * The template object from PHPWord
     * @var PHPWord_Template
     */
    private $template;

    /**
     * The constructor
     *
     * @param Word_Base_Base $word The word object
     * @param string $templatePath The file and path of the template
     */
    public function __construct(Word_Base_Base $word, $templatePath) {
        // store the word object
        $this->word = $word;

        // read the template
        $this->template = $this->readTemplate($templatePath);
    }

    /**
     * Reads the template file in
     *
     * @param string $templatePath The file and path of the template file to
     * read in
     * @return PHPWord_Template
     */
    private function readTemplate($templatePath) {
        return $this->word->getObject()->loadTemplate($templatePath);
    }

    /**
     * Saves the template to the filesystem
     */
    protected function save() {
        $this->template->save($this->word->getPath().$this->word->getFilename());
    }

    /**
     * Replaces a tag with in the template (of the format ${tag_name}) with
     * the value. Note that this function is not idempotent, so after being
     * run once, cannot be run on the same tag again
     *
     * @param string $tag The tag to replace with the value
     * @param string $value The value to replace the tag name with
     */
    protected function setTag($tag, $value) {
        $value = is_array($value) ? array($this->cleanValue($value[0])) : $this->cleanValue($value);
        $this->template->setValue($tag, $value);
    }

    /**
     * Cleans the values before entry to strip any extra tags and ampersands
     *
     * @param string $value The value to clean up
     * @return string The cleaned value
     */
    protected function cleanValue($value) {
        return str_replace('&', 'and', strip_tags($value));
    }

    /**
     * This finds a tag in a table and duplicates the found row x number of
     * times, then replaces the data with that found. For instance if you
     * create a table such as:
     * First name    | Last name
     * ${firstname} | ${lastname}
     *
     * Then with a tag of ${first_name} and an n*m array this method will
     * duplicate the rows n times and add the data from m into these rows. This
     * method is not idempotent so after being run once cannot be run on the
     * same tag again.
     *
     * If this function fails, check the template xml to make sure your tags
     * have not been split over multiple xml tags.
     *
     * @param string $tag A tag in a table cell, in the row to duplicate
     * @param array $values An array of array of values to fit into the table
     */
    protected function createRows($tag, array $values) {
        // clone the template
        $this->template->cloneRow($tag, sizeof($values));

        // update the template with the new values
        foreach ($values as $key => $data) {
            foreach ($data as $rowTag => $rowValue) {
                $this->template->setValue($rowTag.'#'.($key + 1), $rowValue);
            }
        }
    }

    /**
     * Removes a row from a table by finding a tag in a cell and stripping
     * the row out
     *
     * @param string $tag A tag in a table cell in the row to remove
     */
    protected function removeRow($tag) {
        // get the xml
        $xml = $this->template->getXML();

        // find the tag
        $pos = strpos($xml, '${'.$tag.'}');

        // strip out the tr xml tag
        $startPos = strrpos(substr($xml, 0, $pos), '<w:tr ');
        $endPos = strpos($xml, '</w:tr>', $pos) + 7;
        if ($pos !== false && $startPos !== false && $endPos !== false) {
            $xml = substr($xml, 0, $startPos).substr($xml, $endPos);
        }

        // update the xml
        $this->template->setXML($xml);
    }

    /**
     * Removes a column from a table based on the tag.
     *
     * If you want the table to always be 100 percent, you can change the
     * autofit settings for the table in word, or manually alter the
     * <w:tblW property in the XML, using the options specified at
     * http://www.schemacentral.com/sc/ooxml/t-w_ST_TblWidth.html
     *
     * The auto-width will not look right in Wordpad, only Word 2007 and above.
     *
     * @param string $tag The tag name used to find the column to remove
     */
    protected function removeColumn($tag) {
        // get the xml
        $xml = $this->template->getXML();

        // find the tag
        $pos = strpos($xml, '${'.$tag.'}');

        // find the position of the cell in the row
        $rowStartPos = strrpos(substr($xml, 0, $pos), '<w:tr ');
        $cellNum = substr_count(substr($xml, $rowStartPos, $pos - $rowStartPos), '<w:tc>') - 1;
        if ($pos === false || $rowStartPos === false) {
            return;
        }

        // find the start of the table
        $tableStartPos = strrpos(substr($xml, 0, $pos), '<w:tbl>');
        $tableEndPos = strpos($xml, '</w:tbl>', $pos) + 8;
        $tableCode = substr($xml, $tableStartPos, $tableEndPos - $tableStartPos);
        if ($tableStartPos === false || $tableEndPos === false) {
            return;
        }

        // find each row in reverse order <w:tr
        preg_match_all('#\<w\:tr.*\<\/w\:tr\>#Usi', $tableCode, $rowMatches, PREG_OFFSET_CAPTURE | PREG_SET_ORDER);
        $rowMatches = array_reverse($rowMatches);

        // remove the correct cell from each row
        foreach ($rowMatches as $match) {
            // find the correct cell
            $rowCode = $match[0][0];
            preg_match_all('#<w:tc>.*\<\/w\:tc\>#Usi', $rowCode, $cellMatches, PREG_OFFSET_CAPTURE | PREG_SET_ORDER);

            // remove the cell from the xml code
            if (isset($cellMatches[$cellNum])) {
                $cellStartPos = $tableStartPos + $match[0][1] + $cellMatches[$cellNum][0][1];
                $cellEndPos = $cellStartPos + strlen($cellMatches[$cellNum][0][0]);
                $xml = substr($xml, 0, $cellStartPos).substr($xml, $cellEndPos);
            }
        }

        // update the xml
        $this->template->setXML($xml);
    }

    /**
     * Clones a paragraph and adds data to it
     *
     * @param string $tag The tag in the paragraph to clone
     * @param array $data An array of data to replace into the paragraphs
     */
    protected function cloneParagraph($tag, array $data) {
        // get the xml
        $xml = $this->template->getXML();

        // find the tag
        $pos = strpos($xml, '${'.$tag.'}');

        // find the start of the row
        $startPos = strrpos(substr($xml, 0, $pos), '<w:p ');
        $endPos = strpos($xml, '</w:p>', $pos) + 6;

        // loop through the data, building a temp string to insert
        $result = '';
        $code = substr($xml, $startPos, $endPos - $startPos);
        foreach ($data as $row) {
            $result .= str_replace('${'.$tag.'}', $this->cleanValue($row), $code);
        }

        // copy the code to follow this
        $xml = substr($xml, 0, $startPos).$result.substr($xml, $endPos);

        // update the xml
        $this->template->setXML($xml);
    }

    /**
     * Run this once on your template after each change to fix issue such as
     * where the spellchecker has left tags in the xml.
     *
     * Instead of running this every time you output, it's worth running this
     * once with no data and then saving the repaired output.
     *
     * Note that this won't repair issues where the tag has been altered or
     * not typed in one go.
     */
    public function repairTemplate() {
        // get the xml to repair
        $xml = $this->template->getXML();

        // find the start of tags
        preg_match_all('#\<w\:t\>\$\{#Usi', $xml, $matches, PREG_OFFSET_CAPTURE | PREG_SET_ORDER);

        // loop through them and remove any tags between them. This is in
        // reverse order so any changes won't affect the next match
        $matches = array_reverse($matches);
        foreach ($matches as $match) {
            // find the end tag
            $startPos = $match[0][1];
            $endPos = strpos($xml, '}</w:t>', $startPos);

            // find the tag from the code
            $tagCode = substr($xml, $startPos, $endPos - $startPos + 7);
            $simpleTagCode = '<w:t>'.preg_replace('#<.*>#Usi', '', $tagCode).'</w:t>';

            // add the tag back into the code
            $xml = substr($xml, 0, $startPos).$simpleTagCode.substr($xml, $endPos + 7);
        }

        // save the repaired xml
        $this->template->setXML($xml);
    }
}

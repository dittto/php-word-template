<?php
/**
 * Class Word_Quote_Template
 *
 * A simple test template
 */
class Word_TestTemplate extends Word_Base_Template {
    /**
     * Creates a docx file
     */
    public function createFile() {
        // temporarily repair the file every time during development
        $this->repairTemplate();

        // create a test table
        $data = array();
        $data[] = array('first_name' => 'First', 'last_name' => 'Name');
        $data[] = array('first_name' => 'Another', 'last_name' => 'Name');
        $this->createRows('first_name', $data);

        // save the template
        $this->save();
    }
}
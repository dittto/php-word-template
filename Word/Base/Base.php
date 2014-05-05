<?php

/**
 * Class Word_Base_Base
 *
 * Handles creating .docx using PHP and PHPWord. This uses the patched version
 * from http://jeroen.is/phpword-templates-with-repeating-rows/
 */
abstract class Word_Base_Base
{
    /**
     * The filename to output the docx file as
     * @var string
     */
    private $filename = '';

    /**
     * The path to save the temp docx files to. This needs to be relative to
     * the web dir
     * @var string
     */
    private $path = '../user-files/';

    /**
     * The PHPWord object to use to generate the docx download
     * @var PHPWord
     */
    private $object;

    /**
     * The constructor
     *
     * @param string $filename The filename for the docx file
     * @param string $path The local path to store the file locally
     */
    public function __construct($filename, $path = null) {
        // init vars
        $this->filename = $filename;
        $this->path = $path ? $path : $this->path;

        // init the PHPWord object
        include('../vendor/PHPWord/PHPWord.php');
        $this->object = new PHPWord();
    }

    /**
     * Return the PHPWord object - useful for debugging
     *
     * @return PHPWord The PHPWord object
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * Gets the headers for outputting through Kohana
     *
     * @return string[] An array of the headers required to output the docx file
     */
    public function getHeaders()
    {
        // add the headers
        $headers = array();
        $headers['content-type'] = $this->getMime();
        $headers['Pragma'] = 'cache';
        $headers['Cache-Control'] = 'max-age=0';
        $headers['Content-Disposition'] = 'attachment; filename='.$this->getPath().$this->getFilename();

        return $headers;
    }

    /**
     * We only output in Word 2007 format. As a result this will only return the mime
     * type associated with docx files
     *
     * @return string The mime type for .docx files
     */
    public function getMime()
    {
        return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    }

    /**
     * Gets the temp path for saving the docx files between creating them and outputting them
     *
     * @return string The path to save the temp files to
     */
    public function getPath()
    {
        return $this->path;
    }

    /**
     * Returns the filename of the file to create, and to return back to the user as
     *
     * @return string The filename
     */
    public function getFilename()
    {
        return $this->filename.'.docx';
    }
}
<?php

namespace Alshabalin\DocxTemplate;

use ZipArchive;
use Exception;

class DocxTemplate extends ZipArchive
{
    const TPL_TAG_OPEN  = '{';
    const TPL_TAG_CLOSE = '}';

    /**
     * @var string Temporary directory for docx file
     */
    protected $tempDir;

    /**
     * @var string Temporary docx file
     */
    protected $tempFilename;

    /**
     * @var array
     */
    protected $parts;

    /**
     * Opens docx file
     */
    public function open($docx_file, $flags = NULL)
    {
        if ( ! is_file($docx_file))
        {
            throw new Exception('File '  . $docx_file . ' not found');
        }

        $this->tempFilename = tempnam($this->getTempDir(), 'docx');

        if ( ! copy($docx_file, $this->tempFilename))
        {
            throw new Exception('Cannot copy file '  . $docx_file . ' to temporary directory');
        }

        if (parent::open($this->tempFilename) !== true)
        {
            unlink($this->tempFilename);
            throw new Exception('Unable to unpack file' .  $docx_file);
        }

        return $this;
    }

    /**
     * Replaces variables with the replacement string in document, header1, header2
     */
    public function setData(array $variables, $fill_all_gaps = false)
    {
        if ($fill_all_gaps)
        {
            $all_vars = array_merge(
                $this->readAllVariables('document'),
                $this->readAllVariables('header1'),
                $this->readAllVariables('header2'),
                $this->readAllVariables('footer1'),
                $this->readAllVariables('footer2')
            );

            $variables += $all_vars;
        }

        $keys = array_keys($variables);

        $values = array_values($variables);

        $keys = array_map(function($key){ return static::TPL_TAG_OPEN . $key . static::TPL_TAG_CLOSE; }, $keys);

        $this->fillPart('document', $keys, $values);
        $this->fillPart('header1',  $keys, $values);
        $this->fillPart('header2',  $keys, $values);
        $this->fillPart('footer1',  $keys, $values);
        $this->fillPart('footer2',  $keys, $values);

        return $this;
    }

    public function getPart($section)
    {
        if ( ! isset($this->parts[$section]))
        {
            $this->parts[$section] = $this->cleanPlaceholders($this->getFromName('word/' . $section . '.xml'));
        }

        return $this->parts[$section];
    }

    public function setPart($section, $contents)
    {
        if ($contents)
        {
            $this->addFromString('word/' . $section . '.xml', $contents);
        }

        $this->parts[$section] = $contents;

        return $this;
    }

    public function fillPart($section, $keys, $values)
    {
        $this->setPart($section,  str_replace($keys, $values, $this->getPart($section)));
    }

    public function readAllVariables($section)
    {
        $part = $this->getPart($section);

        if (preg_match_all('/' . preg_quote(static::TPL_TAG_OPEN) . '(.+?)' . preg_quote(static::TPL_TAG_CLOSE) . '/s', $part, $matches))
        {
            return array_fill_keys($matches[1], '');
        }

        return [];
    }

    /**
     * Saves file
     */
    public function save($filename)
    {
        $this->close();

        if ( ! rename($this->tempFilename, $filename))
        {
            throw new Exception('Unable to save file ' . $filename);
        }
    }

    /**
     * Sets temporary directory
     *
     * @param string $dir
     * @return $this
     * @throws Exception
     */
    public function setTempDir($dir)
    {
        if ( ! is_dir($dir))
        {
            throw new Exception('Directory ' . $dir . ' not found');
        }

        $this->tempDir = $dir;

        return $this;
    }

    protected function getTempDir()
    {
        if ( ! isset($this->tempDir))
        {
            $this->tempDir = sys_get_temp_dir();
        }

        return $this->tempDir;
    }


    /**
     * Strip tags in placeholders
     *
     * @param $raw_xml
     * @return string
     */
    protected function cleanPlaceholders($raw_xml)
    {
        $raw_xml = preg_replace_callback(
            '/' . preg_quote(static::TPL_TAG_OPEN) . '.+?' . preg_quote(static::TPL_TAG_CLOSE) . '/s',
            function ($match) { return strip_tags($match[0]); },
            (string)$raw_xml
        );

        return $raw_xml;
    }
}
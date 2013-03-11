<?php

/**
 * @file
 * DOCx is a PHP library designed to manipulate the content of .docx files.
 *
 * Maintainer: Stian Hanger (pdnagilum@gmail.com)
 *
 * .docx files (Microsoft Word 2007 and newer) are basically a zipped archive of
 * .xml files which holds different contents. If you unzip a .docx file you will
 * get a lot of files, among them word/document.xml, word/header.xml, and
 * word/footer.xml. For some reason Word adds multiple documents for each section.
 * So there might be 3 footer files, named: word/footer1.xml, word/footer2.xml,
 * and word/footer3.xml. This library takes that into account when manipulating
 * the variables.
 *
 * Functions of the library:
 *
 * * cleanTagVars - Cleans out the excessive '${' and '}' tags in the document.
 * * close - Resets the class for a new run.
 * * load - Loads a .docx file into memory and unzips it.
 * * save - Saves the temporary buffer to disk.
 * * setValue - Does a global search and replace with the given values.
 * * setValues - Does a global search and replace with an array of values.
 * * setValueDocument - Does a search and replace in the 'document' part of the file.
 * * setValueFooter - Does a search and replace in the 'footer' part of the file.
 * * setValueHeader - Does a search and replace in the 'header' part of the file.
 */

/**
 * Define the base path of the library.
 *
 * This will be used as a point of reference for creating temporary files
 */
define('DOCX_BASE_PATH', dirname(__FILE__) . '/');

/**
 * The DOCx class.
 */
class DOCx {
  private $_filename     = NULL;
  private $_filepath     = NULL;
  private $_tempFilename = NULL;
  private $_tempFilepath = NULL;
  private $_zipArchive   = NULL;

  private $_documents    = array();
  private $_footers      = array();
  private $_headers      = array();

  /**
   * Initiate a new instance of the DOCx class.
   *
   * @param string $filepath
   *   The .docx file to load.
   */
  public function __construct($filepath = NULL) {
    if ($filepath !== NULL) {
      $this->load($filepath);
    }
  }

  /**
   * Cleans out the excessive '${' and '}' tags in the document.
   */
  public function cleanTagVars() {
    $this->setValueModifier($this->_documents, '${', '');
    $this->setValueModifier($this->_documents, '}', '');

    $this->setValueModifier($this->_footers, '${', '');
    $this->setValueModifier($this->_footers, '}', '');

    $this->setValueModifier($this->_headers, '${', '');
    $this->setValueModifier($this->_headers, '}', '');
  }

  /**
   * Resets the class for a new run.
   */
  public function close() {
    $this->_filename = NULL;
    $this->_filepath = NULL;
    $this->_tempFilename = NULL;
    $this->_tempFilepath = NULL;
    $this->_zipArchive = NULL;

    $this->_documents = array();
    $this->_footers = array();
    $this->_headers = array();
  }

  /**
   * Compiles a fresh XML document from entries.
   *
   * @param array $entries
   *   The entries to compile from.
   *
   * @return string
   *   Newly formed XML.
   */
  private function compileXML($entries) {
    $output = '';

    if (count($entries)) {
      foreach ($entries as $entry) {
        if (!empty($entry['tag'])) {
          $output .= '<' . $entry['tag'] . '>' . $entry['content'];
        }
      }
    }

    return $output;
  }

  /**
   * Loads a .docx file into memory and unzips it.
   *
   * @param string $filename
   *   The .docx file to load.
   */
  public function load($filepath) {
    $this->_filename = (strpos($filepath, '/') !== FALSE ? substr($filepath, strrpos($filepath, '/') + 1) : $filepath);
    $this->_filepath = $filepath;

    $this->_tempFilename = '.' . time() . '.temp.docx';
    $this->_tempFilepath = DOCX_BASE_PATH . $this->_tempFilename;

    copy(
      $this->_filepath,
      $this->_tempFilepath
    );

    $this->_zipArchive = new ZipArchive();
    $this->_zipArchive->open($this->_tempFilepath);

    $iterators = array(
      '',
    );

    for ($i = 0; $i < 100; $i++) {
      $iterators[] = $i;
    }

    for ($i = 0; $i < count($iterators); $i++) {
      $xmlDocument = $this->_zipArchive->getFromName('word/document' . $iterators[$i] . '.xml');
      $xmlFooter = $this->_zipArchive->getFromName('word/footer' . $iterators[$i] . '.xml');
      $xmlHeader = $this->_zipArchive->getFromName('word/header' . $iterators[$i] . '.xml');

      if (is_string($xmlDocument) && !empty($xmlDocument)) {
        $this->_documents[] = array(
          'content'   => $xmlDocument,
          'localName' => 'word/document' . $iterators[$i] . '.xml',
          'text'      => $this->stripMarkup($xmlDocument),
          'xml'       => $this->splitXML(explode('<', $xmlDocument)),
        );
      }

      if (is_string($xmlFooter) && !empty($xmlFooter)) {
        $this->_footers[] = array(
          'content'   => $xmlFooter,
          'localName' => 'word/footer' . $iterators[$i] . '.xml',
          'text'      => $this->stripMarkup($xmlFooter),
          'xml'       => $this->splitXML(explode('<', $xmlFooter)),
        );
      }

      if (is_string($xmlHeader) && !empty($xmlHeader)) {
        $this->_headers[] = array(
          'content'   => $xmlHeader,
          'localName' => 'word/headers' . $iterators[$i] . '.xml',
          'text'      => $this->stripMarkup($xmlHeader),
          'xml'       => $this->splitXML(explode('<', $xmlHeader)),
        );
      }
    }
  }

  /**
   * Saves the temporary buffer to disk.
   *
   * @param string $filepath
   *   The file to save to. If none is given the temp file is used.
   *
   * @return string
   *   The filepath of the save file.
   */
  public function save($filepath = NULL) {
    if (count($this->_documents)) {
      foreach ($this->_documents as $document) {
        $this->_zipArchive->addFromString($document['localName'], $this->compileXML($document['xml']));
      }
    }

    if (count($this->_footers)) {
      foreach ($this->_footers as $footer) {
        $this->_zipArchive->addFromString($footer['localName'], $this->compileXML($footer['xml']));
      }
    }

    if (count($this->_headers)) {
      foreach ($this->_headers as $header) {
        $this->_zipArchive->addFromString($header['localName'], $this->compileXML($header['xml']));
      }
    }

    $this->_zipArchive->close();

    if ($filepath !== NULL) {
      copy(
        $this->_tempFilepath,
        $filepath
      );

      return $filepath;
    }
    else {
      return $this->_tempFilepath;
    }
  }

  /**
   * Does a global search and replace with the given values.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValue($search, $replace) {
    if (is_string($search) &&
      is_string($replace)) {
      $this->setValueDocument($search, $replace);
      $this->setValueFooter($search, $replace);
      $this->setValueHeader($search, $replace);
    }
  }

  /**
   * Does a global search and replace with an array of values.
   *
   * @param array $values
   *   A keyed array with search and replaces values.
   */
  public function setValues($values) {
    if (is_array($values) &&
      count($values)) {
      foreach ($values as $key => $value) {
        $this->setValue($key, $value);
      }
    }
  }

  /**
   * Does a search and replace in the 'document' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueDocument($search, $replace) {
    if (is_string($search) &&
      is_string($replace)) {
      $this->setValueModifier($this->_documents, $search, $replace);
    }
  }

  /**
   * Does a search and replace in the 'footer' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueFooter($search, $replace) {
    if (is_string($search) &&
      is_string($replace)) {
      $this->setValueModifier($this->_footers, $search, $replace);
    }
  }

  /**
   * Does a search and replace in the 'header' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueHeader($search, $replace) {
    if (is_string($search) &&
      is_string($replace)) {
      $this->setValueModifier($this->_headers, $search, $replace);
    }
  }

  /**
   * Does a search and replace in the given array.
   *
   * @param array $array
   *   The DOCx array to modify.
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   * @param bool $validateSearchVar
   *   Wether or not to validate the search tag as ${TAGNAME}.
   */
  private function setValueModifier(&$array, $search, $replace, $validateSearchVar = TRUE) {
    $cleanSearchTag = '';
    $taggedSearchTag = '';

    if ($validateSearchVar && strlen($search) > 2 && substr($search, 0, 2) !== '${' && substr($search, -1) !== '}') {
      $taggedSearchTag = '${' . $search . '}';
      $cleanSearchTag = $search;
    }
    else {
      $taggedSearchTag = $search;
      $cleanSearchTag = ($validateSearchVar ? substr($search, 2, strlen($search) - 3) : $search);
    }

    if (count($array)) {
      for ($i = 0; $i < count($array); $i++) {
        if (count($array[$i]['xml'])) {
          for ($j = 0; $j < count($array[$i]['xml']); $j++) {
            if (strpos($array[$i]['xml'][$j]['content'], $taggedSearchTag) !== FALSE) {
              $array[$i]['xml'][$j]['content'] = str_replace($taggedSearchTag, $replace, $array[$i]['xml'][$j]['content']);
            }
            elseif ($array[$i]['xml'][$j]['content'] == $cleanSearchTag) {
              $array[$i]['xml'][$j]['content'] = str_replace($cleanSearchTag, $replace, $array[$i]['xml'][$j]['content']);
            }
          }
        }
      }
    }
  }

  /**
   * Splits up XML into separate entities.
   *
   * @param array $lines
   *   The tags of the XML.
   *
   * @return array
   *   Separated XML.
   */
  private function splitXML($lines) {
    $output = array();

    if (count($lines)) {
      foreach ($lines as $line) {
        $tag = '';
        $content = '';
        $temp = '';

        $sp = strpos($line, '>');
        if ($sp !== FALSE) {
          $tag = substr($line, 0, $sp);
          $content = substr($line, $sp + 1);
        }

        $output[] = array(
          'tag'      => $tag,
          'content'  => $content,
          'original' => $content,
        );
      }
    }

    return $output;
  }

  /**
   * Strips away all markup, and returns the text only.
   *
   * @param string $xml
   *   The markuped text.
   *
   * @return string
   *   The stripped text.
   */
  private function stripMarkup($xml) {
    $lines = explode('<', $xml);
    $output = '';

    if (count($lines)) {
      for ($i = 0; $i < count($lines); $i++) {
        $strpos = strpos($lines[$i], ">");

        if ($strpos !== FALSE) {
          $output .= substr($lines[$i], $strpos + 1);
        }
      }
    }

    return $output;
  }
}

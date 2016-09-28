<?php

namespace ImportExport;

use PHPExcel;
use PHPExcel_IOFactory;

class ExcelReader
{
	/**
	 * @var  $_document PHPExcel
	 */
	protected $_document;

	/** @var  $_columnTitles array */
	protected $_columnTitles;

	/** @var  $_callbacks [] Callable */
	protected $_callbacks;


	/**
	 * ExcelReader constructor.
	 *
	 * @param string $filename The file to import.
	 */
	public function __construct($filename)
	{
		$this->_document = PHPExcel_IOFactory::load($filename);
		// convenience method to clean up the column titles
		$this->getColumnTitles('ImportExport\ExcelReader::cleanValues');
	}

	public static function cleanValues($data)
	{
		array_walk($data, function (&$value, $index) {
			$value = strtolower(trim($value));
			$value = strtr($value, array('  ' => ' '));
			$value = preg_replace('/[^\w\d-# ]+/i', '', $value);
			// unspecified column titles default to a letter of the alphabet
			if (!$value) {
				$value = \PHPExcel_Cell::stringFromColumnIndex($index);
			}
		});
		return $data;
	}

	/**
	 * Create a new importer using an uploaded file, if a file was uploaded.
	 *
	 * @return ExcelReader|null
	 */
	public static function createFromUpload()
	{
		if (!isset($_FILES) || empty($_FILES)) {
			return null;
		}
		reset($_FILES);
		$fileFieldName = key($_FILES);
		$tmpFile       = $_FILES[$fileFieldName]['tmp_name'];
		if (is_uploaded_file($tmpFile)) {
			return new self($tmpFile);
		}

		return null;
	}

	/*
	 * Add function to the named callback chain, or invoke the callback chain with the given arguments.
	 * @param string $name The name of the callback chain
	 * @param mixed $args[] Arguments to pass to the callback.
	 *
	 */
	public function event($name, $args)
	{
		if (is_callable($args[0])) {
			$this->_callbacks[$name][] = $args[0];
			return null;
		} else {
			// call previously stored function with provided arguments:
			if (is_array($this->_callbacks[$name])) {
				foreach ($this->_callbacks[$name] as $callback) {
					$args[0] = call_user_func_array($callback, $args);
				}
				return $args[0];
			} else {
				trigger_error("No event handler registered for $name");
				print_r($args);
			}
		}
	}

	public function registerHandler($name, $callable)
	{
		if (is_callable($callable)) {
			$this->_callbacks[$name][] = $callable;
		}
	}

	public function trigger($event, $args)
	{

	}

	public function onReadData($eventHandler)
	{
		return $this->event(__METHOD__, func_get_args());
	}

	/**
	 * @param Callable $eventHandler When the only parameter to this function is a single function reference, it is added to the chain of callbacks invoked by this function
	 *                               whenever the importer needs to determine the internal titles for each of the columns being imported.
	 *
	 * @return null
	 */
	public function getColumnTitles($eventHandler)
	{
		return $this->event(__METHOD__, func_get_args());
	}

	public function parse($sheetIndex = 0)
	{
		$sheet = $this->_document->getSheet($sheetIndex);
		// read the data, skipping any rows that were marked invisible by the creator of the spreadsheet:
		$maxCol = $sheet->getHighestDataColumn();
		$maxRow = $sheet->getHighestDataRow();
		foreach ($sheet->getRowIterator() as $rowIndex => $row) {
			if ($rowIndex > $maxRow) {
				break;
			}
			// reject the record if hidden
			if (!$sheet->getRowDimension($rowIndex)->getVisible()) {
				continue;
			}
			$range   = "A{$rowIndex}:{$maxCol}{$rowIndex}";
			$rowData = current($sheet->rangeToArray($range));
			// read header or data row:
			if (empty($this->_columnTitles)) {
				$this->_columnTitles = $this->getColumnTitles($rowData, $rowIndex);
			} else {
				$rowData = array_combine($this->_columnTitles, $rowData);
				if (!empty(array_filter($rowData))) {
					$this->onReadData($rowData, $rowIndex);
				}
			}
		}
	}
}
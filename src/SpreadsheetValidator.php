<?php

declare(strict_types=1);

namespace Verdient\Hyperf3\SpreadsheetRequest;

use Exception;
use Hyperf\Context\ApplicationContext;
use Hyperf\Contract\TranslatorInterface;
use Hyperf\HttpMessage\Upload\UploadedFile;
use Hyperf\Support\MessageBag;
use Hyperf\Validation\Contract\ValidatorFactoryInterface;
use Hyperf\Validation\ValidationException;
use Verdient\Hyperf3\Validation\Validator;
use Vtiful\Kernel\Excel;

/**
 * 电子表格验证器
 * @author Verdient。
 */
class SpreadsheetValidator extends Validator
{
    /**
     * 文件字段名称
     * @author Verdient。
     */
    protected string $fieldName;

    /**
     * 最少需要的行数
     * @author Verdient。
     */
    protected int $minRows;

    /**
     * 最多允许的行数
     * @author Verdient。
     */
    protected int $maxRows;

    /**
     * 最大允许的文件大小
     * @author Verdient。
     */
    protected int $maxFilesize;

    /**
     * 数据起始行
     * @author Verdient。
     */
    protected int $dataRowStartIndex = 2;

    /**
     * @var string 文件名称
     * @author Verdient。
     */
    protected $fileName;

    /**
     * @var string[] 头部信息
     * @author Verdient。
     */
    protected array $headers = [];

    /**
     * 工作表名称
     * @author Verdient。
     */
    protected string $sheetName;

    /**
     * 数据列的类型
     * @author Verdient。
     */
    protected array $columnTypes = [];

    /**
     * 数据列和字段的映射关系
     * @author Verdient。
     */
    protected array $columnMaps = [];

    /**
     * 当前行号
     * @author Verdient。
     */
    protected int $currentRow = 0;

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function __construct(
        TranslatorInterface $translator,
        array $data,
        array $rules,
        array $messages = [],
        array $customAttributes = [],
        $fieldName = '',
        int $minRows = null,
        int $maxRows = null,
        int $maxFilesize = null,
        int $dataRowStartIndex = 2
    ) {

        $this->initReplacers();
        $this->ensureFallbackMessages();

        if ($dataRowStartIndex < 2) {
            throw new Exception('dataRowStartIndex cannot be less than 2');
        }

        $this->translator = $translator;

        $this->customMessages = $messages;

        $this->customAttributes = $customAttributes;

        $this->fieldName = $fieldName;

        $this->messages = new MessageBag();

        $this->minRows = $minRows;

        $this->maxRows = $maxRows;

        $this->maxFilesize = $maxFilesize;

        $this->dataRowStartIndex = $dataRowStartIndex;

        $this->data = [];

        $this->setRules($rules);

        if (!$this->validateUploadedFile($data)) {
            return;
        }

        $file = $data[$fieldName];

        $this->fileName = $file->getClientFilename();

        if ($excel = $this->prepareSpreadsheet($file)) {
            $this->data = $this->parseSpreadsheetData($excel);
        }
    }

    /**
     * 初始化替换器
     * @author Verdient。
     */
    protected function initReplacers()
    {
        $this->addReplacer('min_rows', function (string $message, string $attribute, string $rule, array $parameters, Validator $validator) {
            return str_replace(':min', $parameters['min'], $message);
        });
        $this->addReplacer('max_rows', function (string $message, string $attribute, string $rule, array $parameters, Validator $validator) {
            return str_replace(':max', $parameters['max'], $message);
        });
        $this->addReplacer('distinct_header', function (string $message, string $attribute, string $rule, array $parameters, Validator $validator) {
            return str_replace(':headers', $parameters['headers'], $message);
        });
        $this->addReplacer('missing_header', function (string $message, string $attribute, string $rule, array $parameters, Validator $validator) {
            return str_replace(':headers', $parameters['headers'], $message);
        });
    }

    /**
     * 确保备用的错误信息已设置
     * @author Verdient。
     */
    protected function ensureFallbackMessages()
    {
        $this->fallbackMessages['min_rows'] = $this->fallbackMessages['min_rows'] ?? 'The :attribute allows up to :max rows except the header row';
        $this->fallbackMessages['max_rows'] =  $this->fallbackMessages['max_rows'] ?? 'The :attribute requires at least :min rows except header row';
        $this->fallbackMessages['distinct_header'] =  $this->fallbackMessages['distinct_header'] ?? 'The :attribute has duplicate headers: :headers';
        $this->fallbackMessages['missing_header'] = $this->fallbackMessages['missing_header'] ?? 'The :attribute missing headers: :headers';
        $this->fallbackMessages['unresolvable'] = $this->fallbackMessages['unresolvable'] ?? 'The :attribute can not be parsed';
        $this->fallbackMessages['multiple_sheets'] = $this->fallbackMessages['multiple_sheets'] ?? 'The :attribute cannot contain multiple sheets';
        $this->fallbackMessages['unix_timestamp'] =  $this->fallbackMessages['unix_timestamp'] ?? 'The :attribute must be a date, datetime, or unix timestamp';
    }

    /**
     * 文件校验规则
     * @return array
     * @author Verdient。
     */
    protected function fileRules()
    {
        $rules = [
            'required',
            'file',
            ['mimes', 'xlsx', 'xls', 'csv'],
            ['mimetypes', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel', 'text/csv', 'text/plain', 'application/octet-stream']
        ];
        if ($this->maxFilesize > 0) {
            $rules[] = ['max', strval($this->maxFilesize)];
        }
        return $rules;
    }

    /**
     * 验证上传的文件
     * @param array $data 请求的数据
     * @return bool
     * @author Verdient。
     */
    protected function validateUploadedFile($data): bool
    {
        /** @var ValidatorFactoryInterface */
        $factory = ApplicationContext::getContainer()
            ->get(ValidatorFactoryInterface::class);
        $validator = $factory->make(
            $data,
            [$this->fieldName => $this->fileRules()],
        );
        if ($validator->fails()) {
            $this->messages = $validator->errors();
            $this->failedRules = $validator->failed();
            return false;
        }
        return true;
    }

    /**
     * 准备表格
     * @param UploadedFile $file 上传的文件
     * @return Excel|false
     * @author Verdient。
     */
    protected function prepareSpreadsheet(UploadedFile $file): Excel|false
    {
        $excel = new Excel([
            'path' => dirname($file->getPathname())
        ]);

        try {
            $spreadsheet = $excel->openFile($file->getFilename());
        } catch (\Throwable $e) {
            $this->addFailure($file->getClientFilename(), 'unresolvable');
            return false;
        }

        $sheetNames = $spreadsheet->sheetList();

        if (count($sheetNames) > 1) {
            $this->addFailure($file->getClientFilename(), 'multiple_sheets');
            return false;
        }

        $sheetName = reset($sheetNames);

        $this->sheetName = $sheetName;

        $sheet = $spreadsheet->openSheet($sheetName, Excel::SKIP_EMPTY_ROW);

        $rowData = $sheet->nextRow();

        $requiredHeaders = $this->getRequiredHeaders();

        if (empty($rowData)) {
            $this->addFailure($file->getClientFilename(), 'missing_header', ['headers' => implode(', ', $requiredHeaders)]);
            return false;
        }

        $sheet = $spreadsheet->openSheet($sheetName, Excel::SKIP_EMPTY_ROW);

        $types = array_fill(0, count($rowData), Excel::TYPE_STRING);

        $rowData = $sheet->nextRow($types);

        $headers = [];

        foreach ($rowData as $index => $value) {
            if (empty($value)) {
                continue;
            }
            if (in_array($value, $requiredHeaders)) {
                $headers[$index] = $value;
            }
        }

        $repeatedHeaders = array_diff_assoc($headers, array_unique($headers));

        if (!empty($repeatedHeaders)) {
            $headerPositions = [];
            foreach ($repeatedHeaders as $column => $name) {
                $headerPositions[] = $name . ' @ ' . Excel::stringFromColumnIndex($column) . '1';
            }
            $this->addFailure($file->getClientFilename(), 'distinct_header', ['headers' => implode(', ', $headerPositions)]);
            return false;
        }

        $missingHeaders = array_diff($requiredHeaders, $headers);

        if (!empty($missingHeaders)) {
            $this->addFailure($file->getClientFilename(), 'missing_header', ['headers' => implode(', ', $missingHeaders)]);
            return false;
        }

        $this->headers = $headers;

        $customAttributes = array_flip($this->customAttributes);

        foreach ($headers as $index => $header) {
            if (isset($customAttributes[$header])) {
                $attribute = $customAttributes[$header];
            } else {
                $attribute = $header;
            }

            if ($this->hasRule($attribute, 'UnixTimestamp')) {
                $types[$index] = Excel::TYPE_TIMESTAMP;
            }

            $this->columnMaps[$index] = $attribute;
        }

        $this->columnTypes = $types;

        return $excel;
    }


    /**
     * 获取需要的头部名称
     * @return string[]
     * @author Verdient。
     */
    protected function getRequiredHeaders(): array
    {
        $headers = [];

        foreach (array_keys($this->rules) as $attribute) {
            $headers[$attribute] = $attribute;
        }

        foreach ($this->customAttributes as $attribute => $attributeName) {
            if (isset($headers[$attribute])) {
                $headers[$attribute] = $attributeName;
            }
        }

        return array_values($headers);
    }

    /**
     * 解析电子表格数据
     * @param Excel $excel 表格对象
     * @return array
     * @author Verdient。
     */
    protected function parseSpreadsheetData(Excel $excel): array
    {
        $data = [];

        $lineNumber = 2;

        while (($rowData = $excel->nextRow($this->columnTypes)) !== null) {

            if ($lineNumber < $this->dataRowStartIndex) {
                continue;
            }

            if (empty(array_filter($rowData))) {
                break;
            }

            $namedData = [];

            foreach ($rowData as $index => $value) {

                if (!isset($this->columnMaps[$index])) {
                    continue;
                }

                $namedData[$this->columnMaps[$index]] = $value;
            }

            $data[] = $namedData;
        }

        return $data;
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function passes(): bool
    {
        if (!$this->messages->isEmpty()) {
            return false;
        }

        $this->ensureFallbackMessages();

        $rawData = $this->data;

        $dataCount = count($rawData);

        if ($this->minRows > 0 && $dataCount < $this->minRows) {
            $this->addFailure($this->fieldName, 'min_rows', ['min' => (string) $this->minRows]);
            return false;
        }

        if ($this->maxRows > 0 && $dataCount > $this->maxRows) {
            $this->addFailure($this->fileName, 'max_rows', ['max' => (string) $this->maxRows]);
            return false;
        }

        $this->distinctValues = [];
        $this->failedRules = [];
        $this->currentRow = $this->dataRowStartIndex;

        foreach ($rawData as &$row) {
            $this->data = $row;
            foreach ($this->rules as $attribute => $rules) {
                $attribute = str_replace('\.', '->', $attribute);
                foreach ($rules as $rule) {
                    $this->validateAttribute($attribute, $rule);
                    if ($this->messages->has($attribute)) {
                        return false;
                    }
                }
            }
            foreach ($this->after as $after) {
                call_user_func($after);
            }
            $this->currentRow++;
        }

        $this->data = $rawData;

        return $this->messages->isEmpty();
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function getDisplayableAttribute(string $attribute): string
    {
        $attribute = parent::getDisplayableAttribute($attribute);
        $pos = array_search($attribute, $this->headers);
        if ($pos !== false) {
            return $attribute . ' @ ' . Excel::stringFromColumnIndex($pos) . $this->currentRow;
        }
        return $attribute;
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function validateDistinct(string $attribute, $value, array $parameters): bool
    {
        if (!isset($this->distinctValues[$attribute])) {
            $this->distinctValues[$attribute] = [];
        }
        if (in_array($value, $this->distinctValues[$attribute])) {
            return false;
        }
        $this->distinctValues[$attribute][] = $value;
        return true;
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function validated(): array
    {
        if ($this->invalid()) {
            throw new ValidationException($this);
        }
        return $this->data;
    }

    /**
     * 验证是不是时间戳
     * @return bool
     * @author Verdient。
     */
    public function validateUnixTimestamp(string $attribute, $value, array $parameters)
    {
        if (is_int($value)) {
            return $value > 0;
        }
        $value = strval($value);
        return ctype_digit($value) && $value > 0 && substr($value, 0, 1) !== '0';
    }
}

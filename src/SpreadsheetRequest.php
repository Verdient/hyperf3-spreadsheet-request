<?php

declare(strict_types=1);

namespace Verdient\Hyperf3\SpreadsheetRequest;

use Hyperf\Contract\ValidatorInterface;
use Hyperf\Validation\Contract\ValidatesWhenResolved;
use Hyperf\Validation\Request\FormRequest;
use Hyperf\Validation\ValidatorFactory;
use Verdient\Hyperf3\Validation\AsAuthorized;

/**
 * 电子表格请求
 * @author Verdient。
 */
class SpreadsheetRequest extends FormRequest implements ValidatesWhenResolved
{
    use AsAuthorized;

    /**
     * 获取文件名称
     * @return string
     * @author Verdient。
     */
    protected function fileName(): string
    {
        return 'file';
    }

    /**
     * 最少需要的行数
     * @return int
     * @author Verdient。
     */
    protected function minRows(): int
    {
        return 1;
    }

    /**
     * 最多允许的行数
     * @return int
     * @author Verdient。
     */
    protected function maxRows(): int
    {
        return 0;
    }

    /**
     * 最多允许的文件大小
     * @return int
     * @author Verdient。
     */
    protected function maxFilesize(): int
    {
        return 0;
    }

    /**
     * 数据起始行
     * @return int
     * @author Verdient。
     */
    protected function dataRowStartIndex(): int
    {
        return 2;
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    protected function validationData(): array
    {
        return $this->getUploadedFiles();
    }

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function validator(ValidatorFactory $factory): ValidatorInterface
    {
        $factory = clone $factory;
        $factory->resolver(function (...$args) {
            $args[] = $this->fileName();
            $args[] = $this->minRows();
            $args[] = $this->maxRows();
            $args[] = $this->maxFilesize();
            $args[] = $this->dataRowStartIndex();
            return new SpreadsheetValidator(...$args);
        });
        return $factory->make($this->validationData(), $this->rules(), $this->messages(), $this->attributes());
    }
}

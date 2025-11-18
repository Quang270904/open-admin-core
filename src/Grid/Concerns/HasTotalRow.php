<?php

namespace OpenAdminCore\Admin\Grid\Concerns;

use Closure;
use OpenAdminCore\Admin\Grid\Tools\TotalRow;

trait HasTotalRow
{
    /**
     * @var array<mixed>
     */
    protected $totalRowColumns = [];

    /**
     * @param string  $column
     * @param Closure|null $callback
     *
     * @return $this
     */
    public function addTotalRow($column, $callback)
    {
        $this->totalRowColumns[$column] = $callback;

        return $this;
    }

    /**
     * @return \Illuminate\Contracts\View\Factory|\Illuminate\View\View|string
     */
    public function renderTotalRow()
    {
        if (empty($this->totalRowColumns)) {
            return '';
        }

        $query = $this->model()->getQueryBuilder();

        $totalRow = new TotalRow($query, $this->totalRowColumns);

        $totalRow->setGrid($this);

        return $totalRow->render();
    }
}

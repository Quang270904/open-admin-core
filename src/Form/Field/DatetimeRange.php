<?php

namespace OpenAdminCore\Admin\Form\Field;

class DatetimeRange extends DateRange
{
    /**
     * @var string
     */
    protected $format = 'YYYY-MM-DD HH:mm:ss';
    protected $view = 'admin::form.daterange';
}

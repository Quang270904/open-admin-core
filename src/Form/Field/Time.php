<?php

namespace OpenAdminCore\Admin\Form\Field;

use OpenAdminCore\Admin\Facades\Admin;

class Time extends Date
{
    /**
     * @var string
     */
    protected $format = 'HH:mm:ss';

    /**
     * @return \Illuminate\Contracts\Foundation\Application|\Illuminate\Contracts\View\Factory|\Illuminate\Contracts\View\View|\Illuminate\Foundation\Application|string
     */
    public function render()
    {
        $this->prepend('<i class="fa fa-clock-o fa-fw"></i>')
            ->defaultAttribute('style', 'width: 150px !important; flex: 0 0 auto !important;');

        $rendered = parent::render();

        $js = "
            var picker = $('{$this->getElementClassSelector()}').parent().data('DateTimePicker');
            var input = $('{$this->getElementClassSelector()}');
            if (picker) {
                input.on('focus', function() {
                    if (!input.val()) {
                        picker.date(moment('00:00:00', 'HH:mm:ss'));
                        input.val(''); 
                    }
                });
            }
            ";  
        Admin::script($js);

        return $rendered;
    }
}

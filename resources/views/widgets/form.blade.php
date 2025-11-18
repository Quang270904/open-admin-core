<form {!! $attributes !!}>
    <div class="box-body fields-group px-3 pt-3">

        @include('admin::widgets.fields')

    </div>

    @if ($method != 'GET')
        <input type="hidden" name="_token" value="{{ csrf_token() }}">
    @endif
    
    <!-- /.box-body -->
    @if(count($buttons) > 0)
    <div class="row box-footer border-top border-1 border-gray-500" style="background-color: inherit;">
        <div class="col-md-{{$width['label']}}"></div>


        <div class="col-md-{{$width['field']}}">

            @if(in_array('submit', $buttons))
            <div class="btn-group pt-3 pe-2 float-end">
                <button type="submit" class="btn btn-primary">{{ $submitLabel ?? trans('admin.submit') }}</button>
            </div>
                
            @foreach($submitRedirects as $redirect)
                <label class="float-end" style="margin: 20px 10px 0 0;">
                    <input type="checkbox" class="after-submit" name="after-save" value="{{ array_get($redirect, 'value') }}" {{ $default_check == array_get($redirect, 'value') ? 'checked' : '' }}> {{ array_get($redirect, 'label') }}
                </label>
            @endforeach

            @endif

            @if(in_array('reset', $buttons))
            <div class="btn-group pt-3 pe-2" style="float: left !important;margin-left: 10px;">
                <button type="reset"class="btn btn-warning text-white">{{ trans('admin.reset') }}</button>
            </div>
            @endif
        </div>
    </div>
    @endif
</form>
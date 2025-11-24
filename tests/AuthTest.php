<?php
namespace App\Plugins\SyncGraph;
// Copyright (c) Microsoft Corporation. All rights reserved.
// Toshihiko Iwamoto revised the article.
// Licensed under the MIT license.

use Illuminate\Validation\ValidationException;
use Illuminate\Support\Collection;
use Exceedone\Exment\Model\CustomTable;
use Exceedone\Exment\Model\CustomColumn;
use Exceedone\Exment\Model\LoginSetting;
use Exceedone\Exment\Enums\ValueType;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Graph\Generated\Models;
use Microsoft\Graph\Generated\Models\Event;
use Microsoft\Graph\Generated\Models\ItemBody;
use Microsoft\Graph\Generated\Models\BodyType;
use Microsoft\Graph\Generated\Models\DateTimeTimeZone;
use Microsoft\Graph\Generated\Models\Location;
use Microsoft\Graph\Generated\Models\Attendee;
use Microsoft\Graph\Generated\Models\EmailAddress;
use Microsoft\Graph\Generated\Models\AttendeeType;
use Microsoft\Graph\Generated\Models\OnlineMeetingProviderType;
use Microsoft\Graph\Generated\Models\Sensitivity;
use Microsoft\Graph\Generated\Models\CalendarRoleType;
use Microsoft\Graph\Generated\Models\EventType;
use Microsoft\Graph\Generated\Users\UsersRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\UserItemRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Events\EventsRequestBuilderGetQueryParameters;
use Microsoft\Graph\Generated\Users\Item\Events\EventsRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Events\EventsRequestBuilderPostRequestConfiguration;
//use Microsoft\Graph\Generated\Users\Item\Events\EventItemRequestBuilderPatchRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Events\Item\EventItemRequestBuilderPatchRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Events\Item\Accept\AcceptPostRequestBody;
use Microsoft\Graph\Generated\Users\Item\Events\Item\Decline\DeclinePostRequestBody;
use Microsoft\Graph\Generated\Users\Item\Events\Item\TentativelyAccept\TentativelyAcceptPostRequestBody;
use Microsoft\Graph\Generated\Users\Item\CalendarView\Delta\DeltaRequestBuilder;
use Microsoft\Graph\Generated\Users\Item\CalendarView\Delta\DeltaRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\CalendarView\Delta\DeltaGetResponse;
use Microsoft\Kiota\Authentication\Oauth\ClientCredentialContext;
use Exceedone\Exment\Controllers\LoginUserController;
use DateTime;
use Exceedone\Exment\Model\CustomValueAuthoritable;
use Exceedone\Exment\Enums\SystemTableName;

class GraphHelper {
    const DOMAINS = ['asai-archi.com' , 'mpd-llc.jp'];

    private static GraphServiceClient $graphClient;
    private static Models\EventCollectionResponse $userEvents;

    /**
     * Graph APIの初期化
     * @return void
     */
    public static function initializeGraphForAuthorze($email = 'admin@asai-archi.com'): void {
        GraphHelper::$userEvents = new Models\EventCollectionResponse();

        // $clientId = env('GRAPH_CLIENT_ID');
        // $tenantId = env('GRAPH_TENANT_ID');
        // $clientSecret = env('GRAPH_CLIENT_SECRET');

        [$local, $domain] = explode('@', $email, 2);
        // 意図的に「ログイン設定」の[ログイン設定表示名]の値に、ドメインを指定しています。
        $login_setteing = LoginSetting::getOAuthSettings(false)->first(function ($record) use ($domain) {
            return $record->login_view_name == $domain;
        });
        if ( is_nullorempty($login_setteing) ){
            throw '<Kajitori> Not Found Login Setting !!';
        }

        $tenantId = ( $domain === 'asai-archi.com' ) ? env('GRAPH_TENANT_ID_AAR') : env('GRAPH_TENANT_ID_MPD');

        $clientId = $login_setteing->getOption('oauth_client_id');
        $clientSecret = $login_setteing->getOption('oauth_client_secret');


        $scopes = ['https://graph.microsoft.com/.default'];

        $tokenContext = new ClientCredentialContext(
            $tenantId,
            $clientId,
            $clientSecret);
        GraphHelper::$graphClient = new GraphServiceClient($tokenContext, $scopes);
    }

    /**
     * 指定したユーザーのプロファイルを取得する。
     * @param string $user_id
     * @param array $selectQueryParameters = ['id','displayName','mail','userPrincipalName','assignedPlans']
     * @return Models\User
     */
    public static function getUserById(string $user_id,
        array $selectQueryParameters = ['id','displayName','mail','userPrincipalName','assignedPlans']): Models\User {

        $requestConfiguration = new UserItemRequestBuilderGetRequestConfiguration();
        $queryParameters = UserItemRequestBuilderGetRequestConfiguration::createQueryParameters();
        $queryParameters->select = $selectQueryParameters;
        $requestConfiguration->queryParameters = $queryParameters;

        return GraphHelper::$graphClient->users()
            ->byUserId($user_id)
            ->get($requestConfiguration)
            ->wait();
    }

    /**
     * メールアドレスでユーザーを検索する
     * @param string $mail
     * @param array $selectQueryParameters = ['id','displayName','mail','userPrincipalName','assignedPlans']
     * @return Models\UserCollectionResponse
     */
    public static function serachUserforMail(string $mail,
        array $selectQueryParameters = ['id','displayName','mail','userPrincipalName','assignedPlans']) {
        $requestConfiguration = new UsersRequestBuilderGetRequestConfiguration();
        $headers = [
                'ConsistencyLevel' => 'eventual',
            ];
        $requestConfiguration->headers = $headers;

        $queryParameters = UsersRequestBuilderGetRequestConfiguration::createQueryParameters();
        $queryParameters->count = true;
        $queryParameters->search = '"mail:'.$mail.'"';
        $queryParameters->orderby = ["displayName"];
        $queryParameters->select = $selectQueryParameters;
        $requestConfiguration->queryParameters = $queryParameters;

        return GraphHelper::$graphClient->users()
            ->get($requestConfiguration)
            ->wait();
    }

    /**
     * テナント内の全ユーザーを取得する。
     * @param array $selectQueryParameters
     * @return Models\UserCollectionResponse
     */
    public static function getUsers(array $selectQueryParameters = ['id', 'displayName','mail','userPrincipalName','assignedPlans']) {

        $requestConfiguration = new UsersRequestBuilderGetRequestConfiguration();
        $queryParameters = UsersRequestBuilderGetRequestConfiguration::createQueryParameters();
        $queryParameters->select = $selectQueryParameters;
        $requestConfiguration->queryParameters = $queryParameters;

        return GraphHelper::$graphClient->users()
            ->get($requestConfiguration)
            ->wait();
    }

    /**
     * 指定したユーザーの予定を取得する。
     * 
     * @param string $user_id ユーザーのメールアドレス
     * @param string $event_id イベントID 
     * @param $init $graphClientを初期化するか否かのフラグ
     * @return Models\EventCollectionResponse
     */
    public static function getUserEventByEventID(string $user_id, string $event_id = '', $init = false) {
        if ( is_null($user_id) ) { return false; }
		if ( is_nullorempty($event_id) ) { return true; }

        if ($init) GraphHelper::initializeGraphForAuthorze($user_id);

        $events = GraphHelper::$graphClient->users()
            ->byUserId($user_id) 
            ->events()
            ->byEventId( $event_id )
            ->get()->wait();

        return $events;
    }

    /**
     * 指定したユーザーの予定を取得する。
     * @param string $user_id
     * @param $init $graphClientを初期化するか否かのフラグ
     * @return Models\EventCollectionResponse
     */
    public static function getUserEventByICalUId(string $user_id, ?string $iCalUId = null, $init = false): ?Models\EventCollectionResponse {
        if (is_null($user_id)) { return null; }

        if ($init) GraphHelper::initializeGraphForAuthorze($user_id);

        $headers = [
            'Prefer' => 'outlook.timezone="Tokyo Standard Time"',
        ];

        // $queryStr = "iCalUId eq '".$iCalUId."'";
        $queryStr = "iCalUId eq '".htmlspecialchars(trim($iCalUId), ENT_QUOTES, 'UTF-8')."'";

        $configuration = new EventsRequestBuilderGetRequestConfiguration();
        $configuration->header = $headers;
        $configuration->queryParameters = new EventsRequestBuilderGetQueryParameters();
        // Only request specific properties
        $configuration->queryParameters->select = ['id','subject','body','bodyPreview','organizer','attendees','start','end','locations','recurrence','isCancelled','iCalUId'];
        // Filter My Organized Schedule Only
        $configuration->queryParameters->filter = trim($queryStr);
        // Sort by start time, newest first
        $configuration->queryParameters->orderby = ['start/dateTime DESC'];
        // Get at most 100 results
        $configuration->queryParameters->top = 25;

        //\Log::debug('    $queryStr : '. $queryStr );

        try {
            return GraphHelper::$graphClient->users()
                ->byUserId($user_id)
                ->events()
                ->get($configuration)
                ->wait();
        } catch (\Throwable $e) {
            \Log::warning("[SyncGraph] Skip invalid iCalUId for {$user_id}: {$iCalUId}. Error: " . $e->getMessage());
            return null;
        }
    }

    /**
     * 指定したユーザーの予定を取得する。
     * @param string $user_id
     * @param $init $graphClientを初期化するか否かのフラグ
     * @return Models\EventCollectionResponse
     */
    public static function getUserEvents(string $user_id, ?bool $isOrganizer = true, ?string $subject = null, ?string $startDateTime = '', $init = false): ?Models\EventCollectionResponse {
        if (is_null($user_id)) { return null; }
        $startDateTimeUTC = is_nullorempty($startDateTime) ? '' : new \Carbon\Carbon($startDateTime);
        if ($init) GraphHelper::initializeGraphForAuthorze($user_id);

        $headers = [
            'Prefer' => 'outlook.timezone="Tokyo Standard Time"',
        ];
        $formatStr = 'y-m-dTH:i:s.u';
        $queryStr = $isOrganizer ? 'isOrganizer eq true' : '';
        $queryStr = $isOrganizer && $subject ? $queryStr. " and " : $queryStr;
        $queryStr .= $subject ? "subject eq '".htmlspecialchars(trim($subject), ENT_QUOTES, 'UTF-8')."' and start/dateTime eq '".$startDateTimeUTC->setTimezone('UTC')->format('Y-m-d\TH:i:s')."'" : '';
        // $queryStr .= $subject ? "subject eq '".trim($subject)."'" : '';

        $configuration = new EventsRequestBuilderGetRequestConfiguration();
        $configuration->header = $headers;
        $configuration->queryParameters = new EventsRequestBuilderGetQueryParameters();
        // Only request specific properties
        $configuration->queryParameters->select = ['id','subject','body','bodyPreview','organizer','attendees','start','end','locations','recurrence','isCancelled','iCalUId','type','isOrganizer','isAllDay'];
        // Filter My Organized Schedule Only
        $configuration->queryParameters->filter = trim($queryStr);
        // Sort by start time, newest first
        $configuration->queryParameters->orderby = ['start/dateTime DESC'];
        // Get at most 100 results
        $configuration->queryParameters->top = 25;

        // \Log::debug('    $queryStr : '. $queryStr );

        return GraphHelper::$graphClient->users()
            ->byUserId($user_id) 
            ->events()
            ->get($configuration)->wait();
    }

    /**
     * 指定したユーザーの差分予定を取得する
     * 
     * @param string $user_Id
     * @param $init $graphClientを初期化するか否かのフラグ
     */
    public static function getUserEventDelta(string $user_id, $deltaLink, $init = false){
        if (is_null($user_id)) { return null; }

        if ($init) GraphHelper::initializeGraphForAuthorze($user_id);
        
        if ( !is_nullorempty($deltaLink) ) {
            try {
                $requestAdapter = GraphHelper::$graphClient->getRequestAdapter();
                $request = new DeltaRequestBuilder($deltaLink, $requestAdapter);
                return $request->get()->wait();
            } catch ( \Exception $ex ) {
                \Log::warning("DeltaLink invalid for user {$user_id}, fallback to full sync. Error: " . $ex->getMessage());
                return self::getUserEventDelta($user_id, null);
            }
        } else {

            $requestConfiguration = new DeltaRequestBuilderGetRequestConfiguration();

            $queryParameters = DeltaRequestBuilderGetRequestConfiguration::createQueryParameters();
            // 差分取得(delta)では、filterがサポートされない
            // "The following parameters are not supported with change tracking over the 'CalendarView' resource: '$orderby, $filter, $select, $expand, $search'."
            // $queryParameters->select = ['id','subject','body','bodyPreview','organizer','attendees','start','end','locations','recurrence','isCancelled'];
            // $queryParameters->filter = 'isOrganizer eq true';
            $queryParameters->select = ['id','subject','body','bodyPreview','organizer','attendees','start','end','locations','recurrence','isCancelled','iCalUId','type','isOrganizer','isAllDay'];
    
            $headers = [
                // 'Prefer' => 'odata.maxpagesize=30',
                'Prefer' => 'odata.maxpagesize=400',
            ];
            $requestConfiguration->headers = $headers;

            // $queryParameters->startDateTime = date('c', strtotime('-30 day'));
            // $queryParameters->endDateTime = date('c', strtotime('+30 day'));
            $fileter_start = new \Carbon\Carbon('now');
            $fileter_start->subMonthNoOverflow()->day = 1;
            // $fileter_end   = new \Carbon\Carbon('last day of next month');
            $fileter_end = new \Carbon\Carbon('now');
            $fileter_end->addDays(60);
            $queryParameters->startDateTime = $fileter_start->toIso8601String();
            $queryParameters->endDateTime   = $fileter_end->toIso8601String();

            // \Log::debug('    $queryParameters : '. $queryParameters->startDateTime . ' - ' . $queryParameters->endDateTime);

            $requestConfiguration->queryParameters = $queryParameters;
        
            try {
                return GraphHelper::$graphClient->users()
                    ->byUserId($user_id)
                    ->calendarView()
                    ->delta()
                    ->get($requestConfiguration)
                    ->wait();    
            } catch (\Exception $e) {
                \Log::debug('  -> Get User Event Faild: ' . $e->getMessage());
                return null;
            }
        }
    }

    /**
     * 渡されたユーザーリストの予定表を取得する(指定がない場合は全ユーザー)
     * Exchangeプランが割り当てられていないユーザーは予定表が存在しないためスキップされる。
     * @param Models\UserCollectionResponse $users
     * @param bool $delta
     * @return array $userEvents
     */
    // public static function getAllUserEvents(Models\UserCollectionResponse $users = null, bool $delta = false): array {
    public static function getAllUserEvents($user_custom_values = null, bool $delta = false): array {
        if ( is_nullorempty($user_custom_values) ) return [];
        // if ( !is_array($user_custom_values) ){
        //     $user_custom_values = [$user_custom_values];
        // }
        if ( !$user_custom_values instanceof Collection ){
            $user_custom_values = [$user_custom_values];
        }

        // 全ユーザー情報取得
        // MEMO: システム管理者を除くユーザーが望ましいかも
        $all_users = CustomTable::getEloquent(SystemTableName::USER)->getValueModel()->all();

        $userEvents = array();
        $statusArr = array();

        foreach ($user_custom_values as $user) {
            \Log::debug('  user : '. $user->getValue('user_name'));
            GraphHelper::initializeGraphForAuthorze($user->getValue('email'));
            $serviceExist = true;
            // $serviceExist = false;
            // // Exchangeプランが割り当てられているか確認
            // foreach ($user->getAssignedPlans() as $assignedPlan) {
            //     If ($assignedPlan->getService() === 'exchange') {
            //         $serviceExist = $assignedPlan->getcapabilityStatus() === 'Enabled';
            //     }
            // }
            
            // ユーザーのカレンダーの共有設定を取得
            //   freeBuzyRead : 自分の予定が入っている時間を閲覧可能
            //   limitedRead : タイトルと場所を閲覧可能
            //   read : すべての詳細を閲覧可能
            //   write : 編集が可能
            //   none : 共有しない
            // $calendar_permission = Self::getUsersCalendarPermissions($user->getValue('email'), true);
            // $calendar_role = $calendar_permission->getValue()[0]->getRole()->value(); // 文字列として取得

            // サービスプランExchangeを持っている場合のみ実行
            if ($serviceExist) {
                $events = $delta ? GraphHelper::getUserEventDelta($user->getValue('email'), $user->getValue('deltalink')) : GraphHelper::getUserEvents($user->getValue('email'));

                if ( is_nullorempty( $events ) ) {
                    \Log::debug('  -> '. $user->getValue('user_name') . ' Events is Empty.');
                    continue;
                }

                if ( $delta ) {
                    // $delta_token = str_replace('$deltatoken=', '', parse_url($events->getOdataDeltaLink())['query'] );
                    // $user->setValue('delta_token', $delta_token)->saveQuietly();

                    $deltalink = $events->getOdataDeltaLink();
                    // \Log::debug('  deltalink : '. $deltalink);
                    // \Log::debug('  next link : '. $events->getOdataNextLink());
                    $user->setValue('deltalink', $deltalink)->saveQuietly();
                }
                $eventArr = $events->getValue();
                // Exmentユーザーテーブルの取得
                $exmentUserTable = CustomTable::getEloquent('user');
                $exmentUserSerachColumn = CustomColumn::getEloquent('email', $exmentUserTable);
                // 設備テーブルのインスタンスを作成
                $equipmentTable = CustomTable::getEloquent('equipment');
                $equipmentSerachColumn = CustomColumn::getEloquent('name', $equipmentTable);

                foreach ($eventArr as $event) {
                    $raw_organizer = $event->getOrganizer()?->getEmailAddress()->getAddress();

                    // if ( is_nullorempty($event->getStart()) ) {
                    //     \Log::debug('    subject : '. $event->getSubject() . ' / type : ' . $event->getType()->value() . ' / organizer : ' . $raw_organizer);
                    // } else {
                    //     \Log::debug('    start : '. $event->getStart()?->getDateTime() . ' / subject : '. $event->getSubject() . ' / type : ' . $event->getType()->value() . ' / organizer : ' . $raw_organizer);
                    // }
                    \Log::debug('    start : '. $event->getStart()?->getDateTime() . ' / subject : '. $event->getSubject() . ' / type : ' . ($event->getType()?->value() ?? 'unknown') . ' / organizer : ' . $raw_organizer);

                    $statusArr = [];
                    // キャンセル or 削除のチェック
                    if ( isset( $event->getBackingStore()->get('additionalData')['@removed']) ) {
                        // Exmentに追加できる配列に変換
                        $arr = array(
                            'event' =>
                                array(
                                    'event_id' => $event->getId(),
                                    'isCancelled' => true,
                                ),
                            'status' => $statusArr
                        );
                        $userEvents[] = $arr;
                        continue;
                    }

                    $require_attendees = [];
                    $optional_attendees = [];
                    $guest_require_attendees = '';
                    $guest_optional_attendees = '';
                    $is_organizer = Self::isOrganizerEx( $event, $user );
                    $exist_attendees = false;

                    // ------------------------------------------------------------------------------------------------------------------
                    // 開催者によるイベントの取得ができないイベントへの対応
                    //
                    // 対象イベント条件
                    // - イベント取得対象者が開催者でない
                    // - イベントが繰り返しイベントのサブイベント（固有の変更を加えられていない）でない
                    
                    if ( !$is_organizer && $event->getType()->value() !== 'occurrence' ) {
                        // イベント取得対象者が開催者でなく、出席者ナシ　＞　想定外ありえない
                        if ( is_nullorempty($event->getAttendees()) ){
                            \Log::debug('    -- Skip (イベント取得対象者が開催者でなく、出席者ナシ) [ Type : ' . $event->getType()->value() .' ]');
                            continue;
                        }

                        // ---------------------------------------------------------------------
                        // 出席者の中に、イベント取得対象ユーザー($user)が存在するかチェック
                        // 存在する > 
                        //   Case 1 : 全ユーザー内に開催者がいる > 後続の処理スキップ
                        //     ⇒ 開催者のユーザーでイベントを取得できるから
                        //
                        //   Case 2 : 全ユーザー内に開催者がいない > 後続の処理を継続
                        //     ⇒ 開催者のユーザーでイベントを取得できないから
                        //     ⇒ 開催者 : グループ or Teamsチャネル or 外部ユーザー or ★未登録のAAR/MPDユーザー★
                        //
                        // 存在しない > 
                        //   Case 3 : 後続の処理を継続
                        //     ⇒ 開催者のユーザーでイベントを取得できないから
                        //     ⇒ 出席者 : グループ or Teamsチャネル
                        // ---------------------------------------------------------------------
                        foreach ($event->getAttendees() as $attendee) {
                            if ( $attendee->getEmailAddress()->getAddress() == $user->getValue('email') ) {
                                $exist_attendees = true;
                                /*
                                // 開催者がAAR/MPDドメインでない場合＝開催者によるレコード登録がなされない場合、後続処理を継続
                                [$local, $domain] = explode('@', $raw_organizer, 2);
                                if ( in_array( $domain, self::DOMAINS ) ) continue 2;
                                */

                                $organizer_user = $all_users->filter(function ($value) use($raw_organizer) {
                                    return $value->getValue('email') == $raw_organizer;
                                });
                                if ( $organizer_user->count() > 0 ) {
                                    // Case 1
                                    \Log::debug('    -- Skip (主催者ユーザーおる) [ user_name : ' . $organizer_user->first()->getValue('user_name') .' ]');

                                    if ( $event->getType()->value() === 'seriesMaster' ) {
                                        \Log::debug('      -- type = seriesMaster なので、特定の情報のみ保存');
                                        $arr = array(
                                            'event' =>
                                                array(
                                                    'event_id' => $event->getId(),
                                                    'raw_organizer' => $raw_organizer,
                                                    'skip' => true,
                                                ),
                                            'status' => $statusArr
                                        );
                                        $userEvents[] = $arr;
                                    }
                                    continue 2;
                                }
                                // Case 2
                                break;
                            }
                        }

                        // Case 3
                        if ( !$exist_attendees ){
                            // 必須出席者に追加
                            //   ⇒追加しないとカレンダーに表示できない
                            $require_attendees []= (string)$user->id;

                            // 出欠回答テーブルに追加
                            //   ⇒追加しないと出欠回答ができない
                            $statusArr []= array(
                                'attendee' => $user->getValue('user_name') . '<' . $user->getValue('email') . '>',
                                'require' => 1,
                                'status' => '0',
                                'event_id' => $event->getId(),
                                'series_master_event_id' => $event->getSeriesMasterId() ?? null,
                                'is_group_user' => true
                            );

                            \Log::debug('    開催者にも出席者にも '.$user->getValue('email').' が存在しない [開催者:'. $raw_organizer .' / タイトル:' .$event->getSubject(). ']');
                            \Log::debug('    iCalUId : '.$event->getICalUId());
                        }
                    }

                    $start = date_create($event->getStart()->getDateTime(), timezone_open($event->getStart()->getTimeZone()));
                    // UTC時間を保存 = Eventの検索に必要
                    $startUTCTime = date_create($event->getStart()->getDateTime(), timezone_open('UTC'));
                    date_timezone_set($start, timezone_open('Asia/Tokyo'));
                    $end = date_create($event->getEnd()->getDateTime(), timezone_open($event->getEnd()->getTimeZone()));
                    date_timezone_set($end, timezone_open('Asia/Tokyo'));

                    // 終日対応
                    if ( $event->getIsAllDay() ){
                        $start->setTime(9,0,0);
                        $end->modify('-1 day')->setTime(23,0,0);
                    }

                    if ( is_nullorempty($event->getAttendees()) ) $event->setAttendees([]);

                    // 必須の出席者と任意の出席者の振り分けとステータステーブル用配列の作成
                    foreach ($event->getAttendees() as $attendee) {
                        $statusResponse = $attendee->getStatus()->getResponse()->value();
                        $name = $attendee->getEmailAddress()->getName();
                        $mail = $attendee->getEmailAddress()->getAddress();

                        // 1.繰り返しイベントの[series_master_id]がない
                        // 2.主催者でない
                        // 3.$eventに recurrence(繰り返しの定義) がある

                        // 主催者の場合、スキップ
                        if ( $raw_organizer == $mail ) continue;

                        // TODO: EntraIDでUserPrincipalName(UPN)とプライマリーメールアドレスが異なる場合がある
                        // 出席者情報のメールアドレスは　プライマリーメールアドレス　なことに注意する
                        // 運用では　UPN＝プライマリーメールアドレス　とのことだが、絶対ではないと思われる
                        // データ不整合が起きた場合は、emailとemail2の両方で比較する
                        $exmentUser = $exmentUserTable->getValueModel()
                            ->where($exmentUserSerachColumn->getIndexColumnName(), '=', $mail)
                            ->first();

                        $exmentUserId = null;
                        $personal_EntraId = null;
                        $personal_event_id = null;

                        if (isset($exmentUser['id'])) {
                            $exmentUserId = $exmentUser->id;
                            // $personal_EntraId = $exmentUser->getValue('user_code');
                            $personal_EntraId = $mail;
                        }

                        if (!is_null($personal_EntraId)) {
                            $personal_event_result = GraphHelper::getUserEventByICalUId(
                                trim($personal_EntraId),
                                $event->getICalUId(),
                                true
                            );

                            if (is_null($personal_event_result)) {
                                \Log::warning("[SyncGraph] No result from getUserEventByICalUId for user {$personal_EntraId} iCalUId {$event->getICalUId()}");
                                $personal_event = [];
                            } else {
                                $personal_event = $personal_event_result->getValue();
                            }

                            if (!empty($personal_event)) {
                                if (!is_array($personal_event)) {
                                    $personal_event = $personal_event->getValue() ?? [];
                                }
                                if (!empty($personal_event)) {
                                    $personal_event_id = $personal_event[0]->getId();
                                }
                            }

                            // 出席者のイベントIDが取得できない場合(==繰り返しイベントのマスター以外のイベント)
                            if (is_nullorempty($personal_event_id) && $user->getValue('email') == $personal_EntraId = $mail) {
                                $personal_event_id = $event->getId();
                            }
                        }


                        $statusArr []= array(
                            'attendee' => $name . '<' . $mail . '>',
                            'require' => $attendee->getType()->value() === 'required' ? 1 : 0,
                            'status' => (int)str_replace(['none','accepted','tentativelyAccepted','declined'], ['0','1','2','3'], $statusResponse),
                            'event_id' => $personal_event_id,
                            //'original_event_id' => $event->getId(),
                            'series_master_event_id' => $event->getSeriesMasterId() ?? null,
                            'is_group_user' => false
                        );

                        // 必須出席者
                        if ($attendee->getType()->value() === 'required'){
                            // Exmentユーザーの場合はIDを取得
                            if (isset($exmentUserId)) {
                                $require_attendees []= (string)$exmentUserId;
                            } else {
                                $guest_require_attendees .= $name . '<' . $mail . '>,';
                            }
                        }
                        // 任意出席者
                        else {
                            // Exmentユーザーの場合はIDを取得
                            if (isset($exmentUserId)) {
                                $optional_attendees []= (string)$exmentUserId;
                            } else {
                                $guest_optional_attendees .= $name . '<' . $mail . '>,';
                            }
                        }
                    }

                    // 出席者のリストから最後のカンマを削除
                    $guest_require_attendees = rtrim($guest_require_attendees, ',');
                    $guest_optional_attendees = rtrim($guest_optional_attendees, ',');

                    // 設備のリストを取得
                    $locations = [];
                    if (!empty($event->getLocations()) || !is_null($event->getLocations())) {
                        foreach ($event->getLocations() as $location) {
                            $equipment = $equipmentTable->getValueModel()
                                ->where($equipmentSerachColumn->getIndexColumnName(), '=', $location->getDisplayName())
                                ->first();
                            if (is_null($equipment) || empty($equipment)) {
                                continue;
                            }
                            $locations []= $equipment->id;
                        }
                    }

                    $organizer = null;
                    if ( !is_nullorempty($event->getOrganizer()) ){
                        $organizer = $exmentUserTable->getValueModel()
                            ->where($exmentUserSerachColumn->getIndexColumnName(), '=', $raw_organizer)
                            ->first();
                    }

                    // Exmentに追加できる配列に変換
                    $arr = array(
                        'event' =>
                            array(
                                // 'event_id' => ( Self::isOrganizerEx($event, $user) ) ? $event->getId() : null,
                                // 一旦セットして、除去の判断はExment登録時(syncExment)で実施
                                'event_id' => $event->getId(),
                                'subject' => $event->getSubject(),
                                // 'bodyPreview' => preg_replace('/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]/', '', $event->getBodyPreview()),
                                // 'body' => $event->getBody()->getContent(), // <- 無駄な改行が多い
                                'body' => $event->getBody() ? preg_replace('/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]/', '', $event->getBody()->getContent()) : null,
                                'organizer' => is_nullorempty($organizer) ? '' : $organizer['id'],
                                'raw_organizer' => $raw_organizer,
                                'require_attendees' => $require_attendees,
                                'optional_attendees' => $optional_attendees,
                                'guest_require_attendees' => $guest_require_attendees,
                                'guest_optional_attendees' => $guest_optional_attendees,
                                'start' => date_format($start, 'Y-m-d').'T'.date_format($start, 'H:i:s.u'),
                                'start_date' => date_format($start, 'Y-m-d'),
                                'start_time' => date_format($start, 'H:i:s'),
                                'end' => date_format($end, 'Y-m-d').'T'.date_format($end, 'H:i:s.u'),
                                'end_date' => date_format($end, 'Y-m-d'),
                                'end_time' => date_format($end, 'H:i:s'),
                                'locations' => $locations ?? null,
                                'isCancelled' => $event->getIsCancelled(),
                                'isOrganizer' => $event->getIsOrganizer(),
                                'sensitivity' => (!is_nullorempty($event->getSensitivity())) ? $event->getSensitivity()->value() : null,
                                // 'sensitivity' => ($calendar_role == 'freeBusyRead') ? 'private' : $event->getSensitivity()->value(),
                                'showAs' => (!is_nullorempty($event->getShowAs())) ? $event->getShowAs()->value() : null,
                                'is_teams' => $event->getIsOnlineMeeting(),
                                'iCalUId' => $event->getICalUId(),
                                // type==occurrence の場合、attendeeがないのでイベント情報に持たせる
                                'series_master_event_id' => $event->getSeriesMasterId() ?? null,
                                'target_user_email' => $user->getValue('email'),
                                'type' => $event->getType()->value(),
                                'recurrence' => Self::makeRecurrenceValue($event, date_format($start, 'H:i:s'),  date_format($end, 'H:i:s')),
                                'isAllday' => $event->getIsAllDay(),
                                'isGroupUser' => !$exist_attendees
                            ),
                        'status' => $statusArr
                    );

                    // イベントがない場合は新規作成、ある場合は追加する。レガシ記法
                    if (empty($userEvents)) {
                        $userEvents = array($arr);
                    } else {
                        array_push($userEvents, $arr);
                    }
                }
            }
        }

        return $userEvents;
        // return array(
        //     'Events' => $userEvents,
        //     'Status' => $statusArr,
        // );
    }

    /**
     * Graph APIで取得したスケジュールをExmentと同期する
     * @param array $events = [event_id, subject, bodyPreview, organizer, organizer_mail, require_attendees, optional_attendees, guest_require_attendees, guest_optional_attendees, start, end, location, isOrganizer]
     * @return string $result
     */
    public static function syncExment(array $events) {
        if (empty($events)) {
            \Log::debug('  [GraphHelper::syncExment] Skip since Events is empty.');
            return 'Events is empty.';  
        }
        \Log::debug('  [GraphHelper::syncExment] start >>>>>>>>>>');

        $exment_schedule = CustomTable::getEloquent('outlook_events');
        $searchColumn = CustomColumn::getEloquent('event_id', $exment_schedule);
        $searchColumn1 = CustomColumn::getEloquent('iCalUId', $exment_schedule);

        $searchCol_start = CustomColumn::getEloquent('start', $exment_schedule);
        $searchCol_end = CustomColumn::getEloquent('end', $exment_schedule);
        $searchCol_raw_organizer = CustomColumn::getEloquent('raw_organizer', $exment_schedule);
        $outlook_searchCol_series_master_event_id = CustomColumn::getEloquent('series_master_event_id', $exment_schedule);

        $exmen_attendance_status = CustomTable::getEloquent('attendance_status');
        $searchCol_event_id = CustomColumn::getEloquent('event_id', $exmen_attendance_status);
        $searchCol_series_master_event_id = CustomColumn::getEloquent('series_master_event_id', $exmen_attendance_status);

        foreach ($events as $event) {
            if ( array_key_exists('skip', $event['event']) && $event['event']['skip'] ) continue;

            //$exment_schedule->refresh();

            $start_date = $event['event']['start_date'] ?? 'N/A';
            $subject = $event['event']['subject'] ?? 'N/A';
            $organizer = $event['event']['raw_organizer'] ?? 'N/A';
            $type = $event['event']['type'] ?? 'unknown';

            \Log::debug("    start : {$start_date} / subject : {$subject} / type : {$type} / organizer : {$organizer}");

            //$eventId = $event['event_id'];
            $event_id = $event['event']['event_id'];

            // 繰り返しイベント対応
            // if ( !is_nullorempty($event['event']['type']) && $event['event']['type'] == 'occurrence' ) {
            if ( isset($event['event']['type']) && in_array($event['event']['type'], ['exception', 'occurrence']) ) {
                if ( $event['event']['type'] == 'occurrence' ){
                    // 固有性のない繰り返しイベント[occurrence]の場合、以下の手順で attendance_status テーブルを検索
                    //   1. 保存済みの繰り返しイベント(サブ)を event_id == $event['event']['event_id] で検索
                    //   2. $events内に必ずいる、該当繰り返しイベントの親(type == seriesMaster)を特定  
                    //   3. 2.で特定したイベントの$event['status']から全ての event_id を取得
                    //   4. 3.で取得した event_id を検索条件に、attendance_status テーブルを検索
                    //   5. 4.の結果に応じて処理を分岐
                    //        該当レコードあり > レコード更新
                    //        該当レコードなし > レコード作成
                    
                    //$master_event = array_values( array_filter($events, fn ($x) => $x['event']['event_id'] == $event['event']['series_master_event_id'] ) );
                    $master_event = [];
                    foreach ($events as $x) {
                        if ( $x['event']['event_id'] == $event['event']['series_master_event_id'] ){
                            $master_event[] = $x;
                        }
                    }

                    if ( count($master_event) == 0 ) continue;

                    // 終日対応
                    if ( $master_event[0]['event']['isAllday'] ?? false ){
                        $end = date_create($event['event']['end_date']);
                        $end->modify('-1 day');
                        $event['event']['end_date'] = date_format($end, 'Y-m-d');
                        $event['event']['end_time'] = '23:00:00';
                    }

                    $records = $exment_schedule->getValueModel()
                        ->Where($searchCol_start->getIndexColumnName(), '=', (string)$event['event']['start_date'].' '.(string)$event['event']['start_time'])
                        ->Where($searchCol_end->getIndexColumnName(), '=', (string)$event['event']['end_date'].' '.(string)$event['event']['end_time'])
                        ->Where($searchCol_raw_organizer->getIndexColumnName(), '=', (string)$master_event[0]['event']['raw_organizer'])
                        ->get();
                    
                    // レコードが存在しない場合
                    if ( !isset($records) || $records->count() == 0 ) {
                        $record = null;
                    }
                    elseif ( $records->count() == 1) {
                        $record = $records[0];
                    }
                    // 対象レコードが複数ある場合 
                    elseif ( $records->count() > 1 ) {
                        $series_master_event_ids = [];
                        foreach ($events as &$e) {
							$e['event']  = $e['event']  ?? [];
							$e['status'] = $e['status'] ?? [];
						}
						unset($e);

						foreach ($events as $e) {
							if (($e['event']['type'] ?? null) !== 'seriesMaster') continue;

							$series_master_event_ids = [];
							foreach ($e['status'] as $s) {
								if (empty($s['event_id'])) continue;
								$series_master_event_ids[] = $s['event_id'];
							}

							if (!empty($event['event']['series_master_event_id']) &&
								in_array($event['event']['series_master_event_id'], $series_master_event_ids)) {
								break;
							}
						}

                        // $child_record = $exmen_attendance_status->getValueModel()
                        //     ->whereInArrayString($searchCol_series_master_event_id->getIndexColumnName(), $series_master_event_ids)
                        //     ->first();
                        $child_records = [];
                        $record = null;
                        foreach ($records as $r) {
                            $child_records = $r->getChildrenValues(CustomTable::getEloquent('attendance_status')->id);
                            foreach( $child_records as $child_record ){
                                $child_record_obj = (object) $child_record;
                                if ( in_array( $child_record->getValue('series_master_event_id') , $series_master_event_ids) ){
                                    $record = $r;
                                    continue 2;
                                }
                            }
                        }
                    }
                } else {
                    // 固有性の　ある　繰り返しイベント[exception]の場合
                    // a. 主催者の場合
                    //   a-1. event_id で検索（occurrence -> exceptionで変わらない）
                    // b. 主催者でない場合
                    //   b-1. 1.で該当レコードがあれば、そちらを更新。なければ、                    
                    //   b-2. 1.で該当レコードがなければ、iCalUId or event_id で検索

                    // 主催者の場合
                    if ( $event['event']['isOrganizer'] ){
                        /*
                        $records = $exment_schedule->getValueModel()
                            ->Where($searchCol_start->getIndexColumnName(), '=', (string)$event['event']['start_date'].' '.(string)$event['event']['start_time'])
                            ->Where($searchCol_end->getIndexColumnName(), '=', (string)$event['event']['end_date'].' '.(string)$event['event']['end_time'])
                            ->Where($searchCol_raw_organizer->getIndexColumnName(), '=', (string)$event['event']['raw_organizer'])
                            ->get();

                        if ( isset($records) && $records->count() > 0 ) {
                            $record = $records[0];
                        } else {
                            // iCalUIdをキーにイベントを検索
                            if ( array_key_exists('iCalUId', $event['event']) && !is_nullorempty($event['event']['iCalUId']) ){
                                $record = $exment_schedule->getValueModel()
                                    ->where($searchColumn->getIndexColumnName(), '=', $event_id)
                                    ->orWhere($searchColumn1->getIndexColumnName(), '=', $event['event']['iCalUId'])
                                    ->first();
                            }
                        }
                        */
                        $record = $exment_schedule->getValueModel()
                            ->where($searchColumn->getIndexColumnName(), '=', $event_id)
                            ->first();

                    }
                    // 主催者じゃない場合
                    else {
                        $child_record = $exmen_attendance_status->getValueModel()
                            ->where($searchCol_event_id->getIndexColumnName(), '=', $event_id)
                            ->first();
                        
                        if ( isset($child_record) ){
                            $record = $exment_schedule->getValueModel( $child_record->parent_id );
                        } else {
                            $record = null;
                        }
                    }

                }
            } else {
                //\Log::debug('    event_id : '. $event_id . ' / iCalUID : ' . $event['event']['iCalUId']);
                if ( !is_nullorempty($event_id) ){
                    if ( array_key_exists('iCalUId', $event['event']) && !is_nullorempty($event['event']['iCalUId']) ){
                        $record = $exment_schedule->getValueModel()
                            ->where($searchColumn->getIndexColumnName(), '=', $event_id)
                            ->orWhere($searchColumn1->getIndexColumnName(), '=', $event['event']['iCalUId'])
                            ->first();
                    } else {
                        $record = $exment_schedule->getValueModel()
                            ->where($searchColumn->getIndexColumnName(), '=', $event_id)
                            ->first();
                    }
                } else {
                    // event_id なし＝繰り返しイベント、ドメイン外からの招待イベントなどの場合
                    if ( array_key_exists('iCalUId', $event['event']) && !is_nullorempty($event['event']['iCalUId']) ){
                        $record = $exment_schedule->getValueModel()
                            ->orWhere($searchColumn1->getIndexColumnName(), '=', $event['event']['iCalUId'])
                            ->first();
                    } else {
                        //TODO: event_idもiCalUIdもないイベントが存在した場合、対応を検討
                        $record = null;
                    }
                }
            }

            // ----------------------------------------------------------
            // outlook_events レコード特定後の処理
            // ----------------------------------------------------------
            try{
                If (isset($record)) {
                    if ( $event['event']['isCancelled'] ) {
                        // Case Delete
                        
                        // 主催者なら
                        if ( !array_key_exists('isOrganizer', $event['event'] )){
                            \Log::debug('    ERROR: 主催者情報なし');
                        } elseif ( $event['event']['isOrganizer'] ){
                            try {
                                $record->delete();
                            } catch (\Exception $e) {
                                \Log::debug('    ERROR: キャンセルされたイベントを削除できませんでした。');
                                \Log::debug('      -> subject : ' . $event['event']['subject'] . ' / start : ' . $event['event']['start_date']);
                            }
                        } else {
                            // 対象ユーザー以外に、Exmentログインユーザーが参加者にいる場合は、削除しない
                            // if ( count($event['event']['require_attendees']) > 0 || count($event['event']['optional_attendees']) > 0 ) continue;
                            // $delete_ok = false;
                            // foreach ($event['status'] as $value) {
                            //     if ( str_contains( $value['attendee'], $event['event']['target_user_email'] ) ) $delete_ok = true;
                            // }

                            $target_user_id = self::getExmentUserByEmail($event['event']['target_user_email'])->id;

                            // outlook_eventsの参加者をメンテ
                            if ( $event['status'][0]['require'] ){
                                $attendees = array_filter($record->getValue('require_attendees', ValueType::PURE_VALUE)->toArray(), function($val) use($target_user_id) {
                                    return $val != $target_user_id;
                                });
                                $record->setValue('require_attendees', $attendees)->save();
                            } else {
                                $attendees = array_filter($record->getValue('optional_attendees', ValueType::PURE_VALUE)->toArray(), function($val) use($target_user_id) {
                                    return $val != $target_user_id;
                                });
                                $record->setValue('optional_attendees', $attendees)->save();
                            }

                            // attendee_statusのリストをメンテ
                            //$record->getChildrenValues(CustomTable::getEloquent('attendance_status')->id)->where('attendee', 'like', '%' . $event['event']['target_user_email'] . '%')->delete();
                            $del_record = $record->getChildrenValues(CustomTable::getEloquent('attendance_status')->id)->filter(function ($item) use($event) {
                                return str_contains($item->getValue('attendee') , $event['event']['target_user_email'] );
                            });
                            if ( !is_nullorempty($del_record) ) $del_record->first()->delete();
                        }
                    } else  {
                        // Case Update

                        // ★★★ 繰り返しイベントの場合(occurrence = 固有の変更や出席者の回答がないイベント) ★★★
                        // 既に登録済み = 
                        if ( !is_nullorempty($event['event']['series_master_event_id']) && $event['event']['type'] == 'occurrence' ){
                            // outlook_eventsを検索
                            $master_record = $exment_schedule->getValueModel()
                                ->where($searchColumn->getIndexColumnName(), '=', $event['event']['series_master_event_id'])
                                ->first();

                            // outlook_eventsに存在しなかった場合、attendance_statusを検索
                            if ( !isset($master_record) ) {

                                $master_child_record = $exmen_attendance_status->getValueModel()
                                    ->where($searchCol_event_id->getIndexColumnName(), '=', $event['event']['series_master_event_id'])
                                    ->first();
                                If (isset($master_child_record)) {
                                    $master_record = $exment_schedule->getValueModel( $master_child_record->parent_id );
                                } else {
                                    // 繰り返しイベント(マスタ)が存在しない＝繰り返しイベントがキャンセルされた
                                    //$master_record = null;
                                    $record->setValidationDestroy(true);
                                    $record->delete();
                                    continue;
                                }
                            }


                            if ( isset($master_record) ){
                                // マスターイベントと同日の場合、後続の処理をスキップ
                                if ( $event['event']['start_date'] == $master_record->getValue('start_date') ) continue;

                                // ★★★これをしないとseriese_masterのデータが最新にならない場合がある
                                $master_record->refresh();

                                $event['event']['subject'] = $master_record->getValue('subject');
                                $event['event']['body'] = $master_record->getValue('body');
                                $event['event']['organizer'] = $master_record->getValue('organizer', ValueType::PURE_VALUE);
                                $event['event']['isOrganizer'] = $master_record->getValue('raw_organizer') == $event['event']['target_user_email'];
                                $event['event']['raw_organizer'] = $master_record->getValue('raw_organizer');
                                $event['event']['require_attendees'] = $master_record->getValue('require_attendees', ValueType::PURE_VALUE)->toArray();
                                $event['event']['optional_attendees'] = $master_record->getValue('optional_attendees', ValueType::PURE_VALUE)->toArray();
                                $event['event']['guest_require_attendees'] = $master_record->getValue('guest_require_attendees');
                                $event['event']['guest_optional_attendees'] = $master_record->getValue('guest_optional_attendees');
                                $event['event']['locations'] = $master_record->getValue('locations', ValueType::PURE_VALUE)->toArray();
                                $event['event']['sensitivity'] = $master_record->getValue('sensitivity');
                                $event['event']['showAs'] = $master_record->getValue('showAs');
                                $event['event']['is_teams'] = $master_record->getValue('is_teams');

                                // マスターイベントの出欠回答をベースに配列を生成
                                $attendance_status_custom_values = $master_record->getChildrenValues(CustomTable::getEloquent('attendance_status')->id);
                                $i = 0;
                                foreach( $attendance_status_custom_values as $val ){
                                    $event['status'][$i]['attendee'] = $val->getValue('attendee');
                                    $event['status'][$i]['require'] = $val->getValue('require');
                                    $event['status'][$i]['status'] = $val->getValue('status');
                                    $event['status'][$i]['is_group_user'] = $val->getValue('is_group_user');
                                    if ( strpos($val->getValue('attendee'), $event['event']['target_user_email'] ) !== false ){
                                        $event['status'][$i]['event_id'] = $event_id;
                                        $event['status'][$i]['series_master_event_id'] = $event['event']['series_master_event_id'];
                                    } else {
                                        $event['status'][$i]['event_id'] = null;
                                        $event['status'][$i]['series_master_event_id'] = null;
                                    }
                                    $i++;
                                }
                            } else {

                            }
                        }

                        // -----------------------------------------------------------------------------------
                        // 繰り返しイベントのマスターイベントに変更があった場合、全ての子イベントを削除
                        //   変更があった＝recurrenceに変更があった
                        //   (※) ただし、マスターイベント以外の繰り返しイベントのIDは、変更されているはずなので、別の削除ロジックで削除される？
                        // -----------------------------------------------------------------------------------
                        if ( $event['event']['type'] === 'seriesMaster' && $record->getValue('recurrence') !== $event['event']['recurrence']){
                            // マスターイベントに出席者テーブルのレコードがある場合
                            if ( isset($event['status']) && count($event['status']) > 0 ) {
                                $status_ids =[];
                                foreach ($event['status'] as $s) {
                                    if ( is_nullorempty($s['event_id']) ) continue;
                                    $status_ids[] = $s['event_id'];
                                }
                                $child_records = $exmen_attendance_status->getValueModel()
                                    ->whereInArrayString($searchCol_series_master_event_id->getIndexColumnName(), $status_ids)
                                    ->get();

                                foreach ($child_records as $r) {
                                    $parrent = $exment_schedule->getValueModel( $r->parent_id );
                                    \Log::debug('    マスターイベントが変更されたためサブイベントを削除 [subject : '.$parrent->getValue('subject').'/start : '.$parrent->getValue('start').']');
                                    try {
                                        $parrent->delete();
                                    } catch (\Exception $e) {
                                        \Log::debug('    ERROR: サブイベントを削除できませんでした。');
                                    }
                                }
                            } else {
                                // マスターイベントに出席者テーブルのレコードがない場合
                                $child_records = $exment_schedule->getValueModel()
                                    ->where($outlook_searchCol_series_master_event_id->getIndexColumnName(), '=', $event_id)
                                    ->get();

                                foreach ($child_records as $r) {
                                    \Log::debug('    マスターイベントが変更されたためサブイベントを削除 [subject : '.$r->getValue('subject').'/start : '.$r->getValue('start').']');
                                    try {
                                        $r->delete();
                                    } catch (\Exception $e) {
                                        \Log::debug('    ERROR: サブイベントを削除できませんでした。');
                                    }
                                }
                            }
                            
                        }

                        $deleteAttendanceStatus = true;
                        // event_id が異なり iCalUId が同一の場合、チャネルorグループで出席者として招集されている会議の場合
                        // require_attendees, optional_attendees に、重複して登録されるユーザーが存在するため、
                        // 既存のユーザー(IDの配列)とOutlookから連携したユーザー(IDの配列)をマージする
                        if ( $record->getValue('event_id') <> $event_id && $record->getValue('iCalUId') == $event['event']['iCalUId'] ){

                            // 登録対象のイベントに [主催者] がセットされていない＝Exmentに未登録　の場合はマージ
                            // 登録対象のイベントに [主催者] がセットされている＝Exmentに登録済み　の場合、出席者の情報を刷新
                            if ( is_nullorempty( $event['event']['organizer']) || $event['event']['isGroupUser'] ){
                                // $event['event']['require_attendees'] = array_unique(array_merge($event['event']['require_attendees'], $record->getValue('require_attendees', ValueType::PURE_VALUE)->toArray()));
                                $deleteAttendanceStatus = false;
                            }
                        }

                        // 主催者じゃない場合、event_idをクリア
                        if ( !$event['event']['isOrganizer'] ) {
                            unset($event['event']['event_id']);
                            unset($event['event']['series_master_event_id']);
                        }

                        // 常に最新情報で上書き
                        // TODO: 常に上書きで問題ないか要確認
                        $event['event']['require_attendees'] = array_unique(array_merge($event['event']['require_attendees'], $record->getValue('require_attendees', ValueType::PURE_VALUE)->toArray()));

                        $record->setValue($event['event'])->save();

                        if ( isset($event['status']) && count($event['status']) > 0 ) {
                            $attendance_status_custom_values = $record->getChildrenValues(CustomTable::getEloquent('attendance_status')->id);
                            GraphHelper::patchAttendanceStatus($attendance_status_custom_values, $event['status'], $record->id, $deleteAttendanceStatus);    
                        }
                    }
                } else {
                    if ( $event['event']['isCancelled'] ) {
                        // Case Delete
                    } else {
                        // Case Create
                        \Log::debug('    Create');
                        if ( is_nullorempty($event['event']['series_master_event_id']) ){
                            $record = $exment_schedule->getvalueModel();
                        } else {
                            // ★★★ 繰り返しイベントの場合 ★★★
                            // outlook_eventsを検索
                            $master_record = $exment_schedule->getValueModel()
                                ->where($searchColumn->getIndexColumnName(), '=', $event['event']['series_master_event_id'])
                                ->first();

                            // outlook_eventsに存在しなかった場合、attendance_statusを検索
                            if ( !isset($master_record) ) {
                                $master_child_record = $exmen_attendance_status->getValueModel()
                                    ->where($searchCol_event_id->getIndexColumnName(), '=', $event['event']['series_master_event_id'])
                                    ->first();
                                
                                If (isset($master_child_record)) {
                                    $master_record = $exment_schedule->getValueModel( $master_child_record->parent_id );

                                } else {
                                    // 繰り返しイベント(マスタ)が存在しない＝繰り返しイベントがキャンセルされた
                                    //$master_record = null;
                                    \Log::debug('    マスターイベントが存在しない！！ [series_master_event_id : '. $event['event']['series_master_event_id'] . ']');
                                    continue;
                                }
                            }

                            if ( isset($master_record) ){
                                // マスターイベントと同日の場合、後続の処理をスキップ
                                if ( $event['event']['start_date'] == $master_record->getValue('start_date') ) continue;

                                $event['event']['subject'] = $master_record->getValue('subject');
                                $event['event']['body'] = $master_record->getValue('body');
                                $event['event']['organizer'] = $master_record->getValue('organizer', ValueType::PURE_VALUE);
                                $event['event']['isOrganizer'] = $master_record->getValue('raw_organizer') == $event['event']['target_user_email'];
                                $event['event']['raw_organizer'] = $master_record->getValue('raw_organizer');
                                $event['event']['require_attendees'] = $master_record->getValue('require_attendees', ValueType::PURE_VALUE)->toArray();
                                $event['event']['optional_attendees'] = $master_record->getValue('optional_attendees', ValueType::PURE_VALUE)->toArray();
                                $event['event']['guest_require_attendees'] = $master_record->getValue('guest_require_attendees');
                                $event['event']['guest_optional_attendees'] = $master_record->getValue('guest_optional_attendees');
                                $event['event']['locations'] = $master_record->getValue('locations', ValueType::PURE_VALUE)->toArray();
                                $event['event']['sensitivity'] = $master_record->getValue('sensitivity');
                                $event['event']['showAs'] = $master_record->getValue('showAs');
                                $event['event']['is_teams'] = $master_record->getValue('is_teams');

                                // マスターイベントの出欠回答をベースに配列を生成
                                $attendance_status_custom_values = $master_record->getChildrenValues(CustomTable::getEloquent('attendance_status')->id);
                                $i = 0;
                                foreach( $attendance_status_custom_values as $val ){
                                    $event['status'][$i]['attendee'] = $val->getValue('attendee');
                                    $event['status'][$i]['require'] = $val->getValue('require');
                                    $event['status'][$i]['status'] = $val->getValue('status');
                                    $event['status'][$i]['is_group_user'] = $val->getValue('is_group_user');
                                    if ( strpos($val->getValue('attendee'), $event['event']['target_user_email'] ) !== false ){
                                        $event['status'][$i]['event_id'] = $event_id;
                                        $event['status'][$i]['series_master_event_id'] = $event['event']['series_master_event_id'];
                                    } else {
                                        $event['status'][$i]['event_id'] = null;
                                        $event['status'][$i]['series_master_event_id'] = null;
                                    }
                                    $i++;
                                }
                            } else {
                                // マスターイベントがExmentに登録されていない場合、Exmentに登録する
                                \Log::debug('    マスターイベントが存在しない！！ [series_master_event_id : '. $event['event']['series_master_event_id'] . ']');
                            }
                        }

                        // 主催者じゃない場合、event_idをクリア
                        if ( !$event['event']['isOrganizer'] ) $event['event']['event_id'] = null;

                        // $record->setValueStrictly($event['event']);
                        // $record->save();
                        // バリデーションを無視するため
                        $record = $exment_schedule->getvalueModel();
                        $record->setValue($event['event'])->save();

                        // 開催者＝null or 開催者<>ログインユーザーの場合、[データ共有(編集権限)]を削除
                        if ( is_nullorempty( $record->getValue('organizer') ) ){
                            CustomValueAuthoritable::deleteValueAuthoritable($record);
                        }

                        GraphHelper::patchAttendanceStatus(array(), $event['status'], $record->id);
                    }
                }
            }
            catch(ValidationException $ex){
                // エラー内容の取得
                \Log::debug($ex->validator->getMessages());
                return $ex->validator->getMessages();
            }
        }

        \Log::debug('  [GraphHelper::syncExment] end   <<<<<<<<<<');
        return 'Sync Success';
    }

    /**
     * CustomValueのデータからGraph APIのリクエストボディを作成する
     * @param CustomValue $customValue
     * @return Event $requestBody
     */
    public static function createRequestBody($customValue): Event {
        // イベント型のインスタンスを作成
        $requestBody = new Event();
        // subject
        $requestBody->setSubject($customValue->getValue('subject'));
        // start, end
        $start = new DateTimeTimeZone();
        $start->setDateTime($customValue->getValue('start'));        
        $start->setTimeZone('Tokyo Standard Time');
        $requestBody->setStart($start);
        $end = new DateTimeTimeZone();
        $end->setDateTime($customValue->getValue('end'));
        $end->setTimeZone('Tokyo Standard Time');
        $requestBody->setEnd($end);
        // body
        $body = new ItemBody();
        $body->setContentType(new BodyType('html'));
        // $body->setContent(nl2br($customValue->getValue('bodyPreview')));
        $body->setContent( $customValue->getValue('body') );
        $requestBody->setBody($body);

        // locations outlook_eventsテーブルのlocationsカラムに設備テーブルを複数選択肢として保存している。
        // 設備テーブルは会議室名を"name"カラムに保存している
        if ($customValue->getValue('locations') != null) {
            $locations = [];
            foreach ($customValue->getValue('locations') as $equipment) {
                $location = new Location();
                $location->setDisplayName($equipment->getValue('name'));
                $locations []= $location;
            }
            $requestBody->setLocations($locations);
        }

        if ( $customValue->getValue('is_teams') ){
            // プロバイダーをTeamsに固定
            $requestBody->setIsOnlineMeeting(true);
            $requestBody->setOnlineMeetingProvider(new OnlineMeetingProviderType('teamsForBusiness'));    
        } else {
            $requestBody->setIsOnlineMeeting(false);
        }

        // attendees
        // Exment User
        $attendeesArray = [];
        if (!empty($customValue->getValue('require_attendees'))) {
            foreach ($customValue->getValue('require_attendees') as $exmentUser) {
                if (empty($exmentUser)) {
                    continue;
                }

                $attendeesArray []= GraphHelper::createAttendee((string)$exmentUser->getValue('name'), $exmentUser->getValue('email'), true);
            }
        }
        if (!empty($customValue->getValue('optional_attendees'))) {
            foreach ($customValue->getValue('optional_attendees') as $exmentUser) {
                if (empty($exmentUser)) {
                    continue;
                }

                $attendeesArray []= GraphHelper::createAttendee((string)$exmentUser->getValue('name'), $exmentUser->getValue('email'), false);
            }
        }
        // Guest User
        if ($customValue->getValue('guest_require_attendees') != null) {
            foreach (explode(',', $customValue->getValue('guest_require_attendees')) as $user) {
                if (empty($user_id)) {
                    continue;
                }
                preg_match('/(?P<name>.*)<(?P<email>.*)>/', $user, $matches);
                
                $attendeesArray []= GraphHelper::createAttendee($matches('name'), $matches('email'), true);
            }
        }
        if ($customValue->getValue('guest_optional_attendees') != null) {
            foreach (explode(',', $customValue->getValue('guest_optional_attendees')) as $user) {
                if (empty($user_id)) {
                    continue;
                }
                preg_match('/(?P<name>.*)<(?P<email>.*)>/', $user, $matches);

                $attendeesArray []= GraphHelper::createAttendee($matches('name'), $matches('email'), false);
            }
        }
        $requestBody->setAttendees($attendeesArray);

        // キャンセル
        $requestBody->setIsCancelled($customValue->getValue('isCancelled'));
        // 公開
        $requestBody->setSensitivity( new Sensitivity($customValue->getValue('sensitivity')) );

        $requestBody->setAllowNewTimeProposals(true);

        /*
        $requestConfiguration = new EventsRequestBuilderPostRequestConfiguration();
        $headers = [
            'Prefer' => 'outlook.timezone="Tokyo Standard Time"',
        ];
        $requestConfiguration->headers = $headers;
        */

        return $requestBody;
    }

    /**
     * Attendee Objectの生成(冗長が激しいので関数として切出す)
     * @param string $userName
     * @param string $userMail
     * @param bool $required
     */
    public static function createAttendee(string $userName, string $userMail, bool $required): Attendee {
        $attendee = new Attendee();
        $emailAddress = new EmailAddress();
        $emailAddress->setAddress($userMail);
        $emailAddress->setName($userName);
        $attendee->setEmailAddress($emailAddress);
        $attendee->setType(new AttendeeType($required ? 'required' : 'optional'));
        return $attendee;
    }

    /**
     * ExmentのデータをGraph APIでPOSTする
     * @param string $user_id
     * @param string $methods = 'create' | 'update'
     * @param string $event_id = ''
     */
    public static function postGraphSchedule(string $email, string $methods, Event $requestBody, string $event_id = ''): string {
        GraphHelper::initializeGraphForAuthorze($email);
        if (empty($email)) {
            return 'Error: User ID is empty.';
        }

        if (empty($methods)) {
            return 'Error: Method is empty.';
        }

        $headers = [
            'Prefer' => 'outlook.timezone="Tokyo Standard Time"',
        ];

        // メトッドにより処理を分岐
        switch ($methods) {
            case 'create':
                $requestConfiguration = new EventsRequestBuilderPostRequestConfiguration();
                $requestConfiguration->headers = $headers;
                $result = GraphHelper::$graphClient->users()
                    ->byUserId($email)
                    ->events()
                    ->post($requestBody, $requestConfiguration)
                    ->wait();
                break;
            case 'update':
                if (empty($event_id) || is_null($event_id)) {
                    return 'Error: Event ID is empty.';
                }
                $requestConfiguration = new EventItemRequestBuilderPatchRequestConfiguration();
                $requestConfiguration->headers = $headers;
                $result = GraphHelper::$graphClient->users()
                    ->byUserId($email)
                    ->events()
                    ->byEventId($event_id)
                    ->patch($requestBody, $requestConfiguration)
                    ->wait();
                break;
            default:
                return 'Error: Method is not found.';
        }

        return $result->getId();
    }

    /**
     * Graph APIで指定したイベントを削除する
     * @param string $user_id
     * @param string $event_id
     * @return string
     */
    public static function deleteGraphSchedule($email, $event_id): string {
        GraphHelper::initializeGraphForAuthorze($email);

        if (empty($email)) {
            return 'Error: User ID is empty.';
        }

        if (empty($event_id)) {
            return 'Error: Event ID is empty.';
        }

        try {
            GraphHelper::$graphClient->users()
                ->byUserId($email)
                ->events()
                ->byEventId($event_id)
                ->delete()
                ->wait();
            $result = 'Delete Event Success';
        } catch (\Exception $e) {
            \Log::debug('Delete Event Faild: ' . $e->getMessage());
            return 'Delete Event Faild: ' . $e->getMessage();
        }

        return $result;
    }

/*
    Exmentで更新した場合
    招待ユーザーの増減がある
    　　増えた場合は子テーブルにステータスを追加
        　　減った場合は子テーブルからステータスを削除
    招待ユーザーの変更がある
        　　変更がある場合は子テーブルのステータスを更新

    Graph側で変更があった場合
    招待ユーザーの増減がある
        　　増えた場合はExmentに追加
        　　減った場合はExmentから削除
    招待ユーザーの変更がある
        　　変更がある場合はExmentのステータスを更新

    両方で変更があった場合  Exment側の変更を優先する
*/

    /**
     * @param string $attendee = userDisplrayName<userEmail>
     * @param int $parent_id
     */
    public static function removeAttendanceStatus(string $attendee, int $parent_id): void {
        $attendanceStatusTable = CustomTable::getEloquent('attendance_status');
        $attendance_status = $attendanceStatusTable->getValueModel()
            ->where('attendee', '=', $attendee)
            ->where('parent_id', '=', $parent_id)
            ->first();
        if (!is_null($attendance_status) && !empty($attendance_status)) {
            $attendance_status->delete();
        }
    }

    /**
     * ExmentUser検索用 Email用
     * @param string $email
     * @return userObject
     */
    public static function getExmentUserByEmail(string $email) {
        $exmentUserTable = CustomTable::getEloquent('user');
        $exmentUserSearchColumn = CustomColumn::getEloquent('email', $exmentUserTable);

        return $exmentUserTable->getValueModel()
            ->where($exmentUserSearchColumn->getIndexColumnName(), '=', $email)
            ->first();
    }

    /**
     * Attendee stringの生成
     * @param $user
     * @return string 
     */
    public static function generateAttendeeString($user, bool $setUserName = true): string {
        if (is_null($user) || empty($user)) {
            return '';
        }
        if ($setUserName) {
            return $user->getValue('user_name') . '<' . $user->getValue('email') . '>';
        } else {
            return '<' . $user->getValue('email') . '>';
        }
    }

    /**
     * イベントを承認する
     * 
     * @param $outlook_events OutlookEventsのCustomValue
     * @return string
     */
    public static function acceptEvent($outlook_events): string {
        // $before_status = $attendee_status->getValue('status');
        // $attendee_status->setValue('status', 1)->save();

        $attendee_status = Self::getUserAttendanceStatus(\Exment::user()->email, $outlook_events, true);
        try {
            $requestBody = new AcceptPostRequestBody();
            $requestBody->setSendResponse(true);

            GraphHelper::$graphClient->users()->byUserId(\Exment::user()->user_code)
                ->events()->byEventId($attendee_status->getValue('event_id'))
                ->accept()->post($requestBody)->wait();
            $attendee_status->setValue('status', 1)->save();
        } catch (\Exception $e) {
            // $attendee_status->setValue('status', $before_status)->save();

            \Log::debug('Failed to accept event: ' . $e->getMessage());
            return '予期せぬエラー！　会議の出席承諾に失敗しました。管理者にお問い合わせください。エラー内容：' . $e->getMessage();
        }
        return '';
    }

    /**
     * イベントを辞退する
     * 
     * @param $outlook_events OutlookEventsのCustomValue
     * @return string
     */
    public static function declineEvent($outlook_events, $is_enforce = false): string {
        // $before_status = $attendee_status->getValue('status');
        // $attendee_status->setValue('status', 3)->setValue('event_id', '')->save();
        /*
        $event_id = $outlook_events->getValue('event_id');

        if ( $is_enforce ){
            GraphHelper::initializeGraphForAuthorze( \Exment::user()->email );
        } else {
            $attendee_status = Self::getLoginUserAttendanceStatus($outlook_events, true);
            $event_id = $attendee_status->getValue('event_id');
        }

        if ( is_nullorempty($event_id) ) return '';
        */

        $attendee_status = Self::getUserAttendanceStatus(\Exment::user()->email, $outlook_events, true);

        try {
            $requestBody = new DeclinePostRequestBody();
            $requestBody->setSendResponse(true);

            GraphHelper::$graphClient->users()->byUserId(\Exment::user()->user_code)
                ->events()->byEventId($attendee_status->getValue('event_id'))
                ->decline()->post($requestBody)->wait();
            
            if ( !$is_enforce ) $attendee_status->setValue('status', 3)->setValue('event_id', '')->save();

        } catch (\Exception $e) {
            // $attendee_status->setValue('status', $before_status)->save();

            \Log::debug('Failed to decline event: ' . $e->getMessage());
            return '予期せぬエラー！　会議の出席辞退に失敗しました。管理者にお問い合わせください。エラー内容：' . $e->getMessage();
        }
        return '';
    }

    /**
     * イベントを仮承諾する
     * 
     * @param $outlook_events OutlookEventsのCustomValue
     * @return string
     */
    public static function tentativelyAcceptEvent($outlook_events): string {
        // $before_status = $attendee_status->getValue('status');
        // $attendee_status->setValue('status', 2)->save();

        $attendee_status = Self::getUserAttendanceStatus(\Exment::user()->email, $outlook_events, true);
        try {
            $requestBody = new TentativelyAcceptPostRequestBody();
            $requestBody->setSendResponse(true);

            GraphHelper::$graphClient->users()->byUserId(\Exment::user()->user_code)
                ->events()->byEventId($attendee_status->getValue('event_id'))
                ->tentativelyAccept()->post($requestBody)->wait();
            $attendee_status->setValue('status', 2)->save();
        } catch (\Exception $e) {
            // $attendee_status->setValue('status', $before_status)->save();
            \Log::debug('Failed to tentatively accept event: ' . $e->getMessage());
            return '予期せぬエラー！　会議の出席仮承諾に失敗しました。管理者にお問い合わせください。エラー内容：' . $e->getMessage();
        }
        return '';
    }

    /**
     * 打合せテーブルのIDからログインユーザーのEntraIDとイベントIDを特定する。
     * @param int $id
     * @return array
     */
    public static function getParsonalIDSet(int $id): array {
        $login_user = \Exment::user();
        $outlookEventsTable = CustomTable::getEloquent('outlook_events');
        $attendance_status = CustomTable::getEloquent('attendance_status');
        $searchUserColumn3 = CustomColumn::getEloquent('attendee', $attendance_status);

        $attendees = [];
        $eventRecord = $outlookEventsTable->getValueModel($id);

        if (!isset($eventRecord['id'])) {
            return [
                'message' => 'この会議に対する詳細情報が登録されていません。管理者にお問い合わせください。',
            ];
        }

        $EntrId = GraphHelper::serachUserforMail($login_user->email)->getValue()[0]->getId();
        $user_status = $attendance_status->getValueModel()
            ->where('parent_id', '=', $id)
            ->where($searchUserColumn3->getIndexColumnName(), '=', $login_user->user_name.'<'.$login_user->email.'>')
            ->first();
        
        if (!isset($user_status['id'])) {
            return [
                'message' => 'この会議に対する詳細情報が登録されていません。管理者にお問い合わせください。',
            ];
        }
        $eventId = $user_status->getValue('event_id');

        return ['entra_id' => $EntrId, 'event_id' => $eventId, 'attendee_status_id' => $user_status['id']];
    }

    /**
     * 出席者のステータスを更新する(差異などは気にしない)
     * @param $attendance_status
     * @param $childRecords
     * @param array $status
     * @param int $parent_id
     * @param boolean $deleteAttendanceStatus : $statusに含まれない$attendance_statusのレコードを削除する
     */
    public static function patchAttendanceStatus($attendance_status_custom_values, array $status, int $parent_id, bool $deleteAttendanceStatus = true): void {
        // 出席者のステータスを持った配列で出席者テーブルを検索して更新
        foreach ($status as $attendeeStatus) {
            if (!isset($attendance_status_custom_values[0])) {
                $newRecord = array(
                    'parent_id' => $parent_id,
                    'parent_type' => 'outlook_events',
                    'value->attendee' => $attendeeStatus['attendee'],
                    'value->require' => $attendeeStatus['require'],
                    'value->status' => $attendeeStatus['status'],
                    'value->event_id' => $attendeeStatus['event_id'],
                    'value->series_master_event_id' => $attendeeStatus['series_master_event_id'],
                    'value->is_group_user' => $attendeeStatus['is_group_user']
                );
                try {
                    getModelName('attendance_status')::create($newRecord);
                } catch (\Exception $ex) {
                    \Log::debug('Record Create Process Faild '. $ex->getMessage());
                }
            } else {
                $childRecord = $attendance_status_custom_values->filter(function ( $value ) use( $attendeeStatus ) {
                    return $value->getValue('attendee') === $attendeeStatus['attendee'];
                })->first();

                // if (empty($childRecord)) { continue; }
                if (empty($childRecord)) {
                    $newRecord = array(
                        'parent_id' => $parent_id,
                        'parent_type' => 'outlook_events',
                        'value->attendee' => $attendeeStatus['attendee'],
                        'value->require' => $attendeeStatus['require'],
                        'value->status' => $attendeeStatus['status'],
                        'value->event_id' => $attendeeStatus['event_id'],
                        'value->series_master_event_id' => $attendeeStatus['series_master_event_id'],
                        'value->is_group_user' => $attendeeStatus['is_group_user']
                    );
                    try {
                        getModelName('attendance_status')::create($newRecord);
                    } catch (\Exception $ex) {
                        \Log::debug('Record Create Process Faild '. $ex->getMessage());
                    }
                } else {
                    try {
                        $childRecord->setValue('require', $attendeeStatus['require']);
                        $childRecord->setValue('status', $attendeeStatus['status']);
                        if ( !is_nullorempty($attendeeStatus['series_master_event_id']) ) $childRecord->setValue('series_master_event_id', $attendeeStatus['series_master_event_id']);
                        if ( !is_nullorempty($attendeeStatus['event_id']) ) $childRecord->setValue('event_id', $attendeeStatus['event_id']);
                        $childRecord->save();
                    } catch (\Exception $ex) {
                        \Log::debug('Record Create Process Faild '. $ex->getMessage());
                    }
                }
            }
        }
        
        if ( ! $deleteAttendanceStatus ) return;

        // Delete attendees that were removed from the event
        if (isset($attendance_status_custom_values[0])) {
            // Get list of emails from new attendee list
            $newAttendeeEmails = array_map(function($attendeeStatus) {
                // Extract email from "Name<email>" format
                preg_match('/<([^>]+)>/', $attendeeStatus['attendee'], $matches);
                return $matches[1] ?? $attendeeStatus['attendee'];
            }, $status);
            
            // Find and delete old records not in new list
            foreach ($attendance_status_custom_values as $oldRecord) {
                // グループユーザーは削除対象外
                if ( $oldRecord->getValue('is_group_user') ) continue;

                $oldAttendee = $oldRecord->getValue('attendee');
                preg_match('/<([^>]+)>/', $oldAttendee, $matches);
                $oldEmail = $matches[1] ?? $oldAttendee;
                
                if (!in_array($oldEmail, $newAttendeeEmails)) {
                    try {
                        \Log::debug('      [DELETE_ATTENDEE] Removing attendee no longer in event: '.$oldAttendee.' from parent_id='.$parent_id);
                        $oldRecord->delete();
                    } catch (\Exception $ex) {
                        \Log::debug('Failed to delete removed attendee: '. $ex->getMessage());
                    }
                }
            }
        }
    }

    /**
     * 特定のユーザーのカレンダーの共有の権限を取得
     * 
     * @param $email
     * @param $init $graphClientを初期化するか否かのフラグ
     * @return Microsoft\Graph\Generated\Models\CalendarPermission
     */
    public static function getUsersCalendarPermissions($email, $init = false) {
        if (is_null($email)) { return null; }

        if ($init) GraphHelper::initializeGraphForAuthorze($email);

        return GraphHelper::$graphClient->users()->byUserId($email)->calendar()->calendarPermissions()->get()->wait();
    }

    /**
     * ログインユーザーが出席者の出欠回答データ取得
     * (既に「辞退」で回答しているデータは対象外)
     *
     * @param CustomValue $outlook_events OutlookテーブルのCustomValue
     * @param bool $get_event_id 出欠回答テーブルにevent_idがなければ取得するフラグ
     * @return CustomValue 出欠回答テーブルのCustomValue
     */
    public static function getLoginUserAttendanceStatus($outlook_events, $get_event_id = false){
        return Self::getUserAttendanceStatus(\Exment::user()->email, $outlook_events, $get_event_id);
    }

    /**
     * ログインユーザーが出席者の出欠回答データ取得
     * (既に「辞退」で回答しているデータは対象外)
     *
     * @param CustomValue $outlook_events OutlookテーブルのCustomValue
     * @param bool $get_event_id 出欠回答テーブルにevent_idがなければ取得するフラグ
     * @return CustomValue 出欠回答テーブルのCustomValue
     */
    public static function getUserAttendanceStatus($email, $outlook_events, $get_event_id = false)
    {
        $attendance_status = $outlook_events->getChildrenValues(CustomTable::getEloquent('attendance_status')->id)->filter(function ($value) use($email) {
            // return str_contains( $value->getValue('attendee'), \Exment::user()->email ) && !is_nullorempty($value->getValue('event_id')) ;
            return str_contains( $value->getValue('attendee'), $email ) && $value->getValue('status') !== '3';
        });
        
        $myValue = $attendance_status->first();

        if ( $get_event_id ){
            // 出欠回答データのイベントIDが未取得の場合は、イベントIDを入手してセット
            if ( is_nullorempty( $myValue->getValue('event_id') ) ){
                // $startDateTime = date('Y-m-d\TH:i:s+0900', strtotime( $outlook_events->getValue('start') ) );
                $start = new DateTime($outlook_events->getValue('start'));
                //$startDateTime = $t->format(DateTime::ATOM);
                // $personal_event = GraphHelper::getUserEvents(trim(\Exment::user()->email), false, $outlook_events->getValue('subject'), $startDateTime, true)->getValue();
                $personal_event = GraphHelper::getUserEvents(trim(\Exment::user()->email), false, $outlook_events->getValue('subject'), date_format($start, 'Y-m-d').'T'.date_format($start, 'H:i:s.u'), true)->getValue();
                // $personal_event = GraphHelper::getUserEventByEventID(trim(\Exment::user()->email), $myValue->getValue('event_id'), true)->getValue();
                //$myValue->setValue('event_id', $personal_event[0]->getId())->save(); // 呼び出し元でステータスと一緒に保存しているのでsave不要
                $myValue->setValue('event_id', $personal_event[0]->getId());
            }
        }
        return $myValue;
    }

    /**
     * 指定したユーザーの予定を取得する。
     * 
     * @param string $user_id
     * @param $init $graphClientを初期化するか否かのフラグ
     * @return Models\EventCollectionResponse
     */
    public static function existsUserEvents(string $user_id, ?string $event_id = '', $init = false) {
        if ( is_null($user_id) ) { return null; }
		if ( is_nullorempty($event_id) ) { return null; }

        if ($init) GraphHelper::initializeGraphForAuthorze($user_id);

        try {
            $events = GraphHelper::$graphClient->users()
                ->byUserId($user_id) 
                ->events()
                ->byEventId( $event_id )
                ->get()->wait();

            // return !empty($events);
            return $events;
            /*
            $eventArr = $events->getValue();
            if ( count($eventArr) > 0 ){
                return true;
            } else {
                return false;
            }
            */
        } catch (\Exception $ex) {
            \Log::debug('  --- Check Event Exists Faild '. $ex->getMessage());
            return null;
        }
    }

    public static function isOrganizerEx( $event, $user ){
        if ( $event->getIsOrganizer() ) return true;
        if ( ( !is_nullorempty($event->getOrganizer()) && $user->getValue('email') == $event->getOrganizer()->getEmailAddress()->getAddress() ) ) return true;
        return false;
    }

    /**
     * 繰り返しイベント関連情報(recurrence)生成
     * 
     * ここで生成の元になる情報が、登録済みのデータと異なることは、イコール繰り返しイベント(サブ)の一括削除が必要。
     * 
     */
    private static function makeRecurrenceValue(Event $event, $start_time, $end_time){
        if ( $event->getType()->value() !== 'seriesMaster' ) return null;

        $value = [];
        $pattern = $event->getRecurrence()->getPattern();
        $range   = $event->getRecurrence()->getRange();

        $value = [
            'base' => [
                'start_time' => $start_time,
                'end_time' => $end_time
            ],
            'pattern' => [
                'dayOfMonth' => $pattern->getDayOfMonth(),
                'daysOfWeek' => ($pattern->getDaysOfWeek()) ? $pattern->getDaysOfWeek()[0]?->value() : null,
                'firstDayOfWeek' => $pattern->getFirstDayOfWeek()->value(),
                'index' => $pattern->getIndex()->value(),
                'interval' => $pattern->getInterval(),
                'month' => $pattern->getMonth(),
                'type' => $pattern->getType()->value(),
            ],
            'range' => [
                'endDate' => $range->getEndDate()->__toString(),
                'numberOfOccurrences' => $range->getNumberOfOccurrences(),
                'recurrenceTimeZone' => $range->getRecurrenceTimeZone(),
                'startDate' => $range->getStartDate()->__toString(),
                'type' => $range->getType()->value(),
            ]
        ];
        return serialize($value);
    }

    /**
     * M365に存在しない、特定のユーザーのOutlookEventsのレコードを削除
     * 
     * 
     */
    public static function cleanOutlookEventsRecords(CustomTable $custom_table, $user, $fileter_start, $fileter_end){
        //  Exment->M365(存在チェック)
        // $fileter_start = new \Carbon\Carbon('now');
        // $fileter_start->subMonthNoOverflow()->day = 1;
        //$fileter_start = new \Carbon\Carbon('first day of last month');

        // $fileter_end   = new \Carbon\Carbon('first day of next month');
        // $fileter_end->endOfMonth();
        //$fileter_end   = new \Carbon\Carbon('last day of next month');

        //$query = $this->custom_table->getValueQuery()->getQuery();
        $query = getModelName('outlook_events')::query();

        $db_table_name = getDBTableName($custom_table);
        $outlook_events = CustomTable::getEloquent($custom_table->table_name);
        $column1 = CustomColumn::getEloquent('require_attendees', $outlook_events)->getIndexColumnName();
        $column2 = CustomColumn::getEloquent('optional_attendees', $outlook_events)->getIndexColumnName();
        $column3 = CustomColumn::getEloquent('organizer', $outlook_events)->getIndexColumnName();
        $target_user_id = $user->id;

        $query->where('value->start', '>=', $fileter_start->toDateString())
            ->where('value->start', '<=', $fileter_end->toDateString())
            ->where(function ($query) use($db_table_name, $column1, $column2, $column3, $target_user_id) {
            $query->whereInArrayString("$db_table_name.$column1", array($target_user_id))
                ->orWhereInArrayString("$db_table_name.$column2", array($target_user_id))
                ->orWhere("$db_table_name.$column3",$target_user_id)
                //->orWhereNotNull('value->locations')
                //->orWhereNotNull('value->equipments')
                ;
            })->orderBy('value->start', 'asc');
        $records = $query->take(config('exment.calendar_max_size_count', 300))->get();
        
        foreach( $records as $record ){
            \Log::debug('  start : '. $record->getValue('start') .' / subject : '. $record->getValue('subject') . ' / raw_orginizer : '. $record->getValue('raw_organizer'));
            
            // $events = GraphHelper::getUserEvents( $login_user->getValue('email'), true, $record->getValue('subject'), $record->getValue('start') );
            // $events = GraphHelper::existsUserEvents( $login_user->getValue('email'), $record->getValue('event_id'), true );

            // 主催者がシステム内におり、イベントIDがセットされている
            if ( !is_nullorempty($record->getValue('organizer') ) && !is_nullorempty($record->getValue('event_id') )) {
                $events = Self::existsUserEvents( $record->getValue('organizer')->getValue('email'), $record->getValue('event_id'), true );
            } else {
                $attendee_status = Self::getUserAttendanceStatus($user->getValue('email'), $record);
                // 該当ユーザーが[必須出席者]、[任意出席者]に含まれるが、出席者に存在しない
                if ( is_nullorempty( $attendee_status ) ) {
                    \Log::debug('    > Not exists attendee_status !! [record id : '. $record->id . ']');

                    // $events = Self::existsUserEvents( $user->getValue('email'), $record->getValue('event_id'), true );
                    $events = Self::getUserEventByICalUId( $user->getValue('email'), $record->getValue('iCalUId'), true);
                } else {
                    // 出席者に存在するが event_id が未登録の場合、後続の処理をスキップ
                    if ( is_nullorempty($attendee_status->getValue('event_id') ) ){
                        \Log::debug('    > Skip since not exists event_id of atendee_status !! [record id : '. $record->id . ']');
                        continue;
                    }
                    $events = Self::existsUserEvents( $user->getValue('email'), $attendee_status->getValue('event_id'), true );
                }
            }

            if ( empty( $events ) ){
                //$record->delete();
                \Log::debug('    > Delete the record since it is   not scheduled   in outlook. !! [record id : '. $record->id . ']');
            } else {
                //$eventArr = $events->getValue();
                //$event = $eventArr[0];
                if ( method_exists($events,'getIsCancelled') && $events->getIsCancelled() ){
                    //$record->delete();
                    \Log::debug('    > Delete the record since it is   Cancelled   in outlook. !! [record id : '. $record->id . ']');
                }
            }
        }

    }
}
?>
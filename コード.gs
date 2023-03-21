//メニューを構築する
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('▶OAuth認証')
      .addItem('認証の実行', 'startoauth')
      .addItem('テストツイート', 'testtweet')
      .addSeparator()
      .addItem('ログアウト', 'reset')
      .addItem('スクリプトプロパティ', 'openCheck')
      .addToUi();
}

const ME_PINNED_TWEET_ID = PropertiesService.getScriptProperties().getProperty('PINNED_TWEET_ID');
const ME_USER_ID = PropertiesService.getScriptProperties().getProperty('USER_ID');
const ME_LIST_ID = PropertiesService.getScriptProperties().getProperty('LIST_ID');
const TARGET_USER_ID = PropertiesService.getScriptProperties().getProperty('TARGET_ID');

//認証用の各種変数
var  apikey = PropertiesService.getScriptProperties().getProperty('API_KEY');
var  apisecret= PropertiesService.getScriptProperties().getProperty('API_SECRET');
var tokenurl = "https://api.twitter.com/oauth/access_token";
var reqtoken = "https://api.twitter.com/oauth/request_token";
var authurl = "https://api.twitter.com/oauth/authorize";
var endpoint = "https://api.twitter.com/1.1/statuses/update.json";  //ツイートをするエンドポイント
var endpoint2 = "https://api.twitter.com/2/tweets";  //v2のツイートするエンドポイント
var appname = "ここにアプリの名称を入れる";  //アプリの名称

//認証実行
function startoauth(){
  //UIを取得する
  var ui = SpreadsheetApp.getUi();
  
  //認証済みかチェックする
  var service = checkOAuth(appname);
  if (!service.hasAccess()) {
    //認証画面を出力
    var output = HtmlService.createHtmlOutputFromFile('template').setHeight(450).setWidth(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ui.showModalDialog(output, 'OAuth1.0認証');
  } else {
    //認証済みなので終了する
    ui.alert("すでに認証済みです。");
  }
}

//認証チェック用関数
function checkOAuth(serviceName) {
  return OAuth1.createService(serviceName)
    .setAccessTokenUrl(tokenurl)
    .setRequestTokenUrl(reqtoken)
    .setAuthorizationUrl(authurl)
    .setConsumerKey(apikey)
    .setConsumerSecret(apisecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties());
}

//認証コールバック
function authCallback(request) {
  var service = checkOAuth(request.parameter.serviceName);
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証が正常に終了しました');
  } else {
    return HtmlService.createHtmlOutput('認証がキャンセルされました');
  }
}

//アクセストークンURLを含んだHTMLを返す関数
function authpage(){
  var service = checkOAuth(appname);
  var authorizationUrl = service.authorize();
  var html = "<center><b><a href='" + authorizationUrl + "' target='_blank' onclick='closeMe();'>アクセス承認</a></b></center>"
  return html;
}

var endpoint2 = "https://api.twitter.com/2/tweets";

//テストツイートする
function testtweet(){
  //message本文
  var message = {
    //テキストメッセージ本文
    text: 'Testtweet from GAS'
  }

  //リクエストオプション
  var options = {
    "method": "post",
    "muteHttpExceptions" : true,
    'contentType': 'application/json',
    'payload': JSON.stringify(message)
  }

  //リクエスト実行
  var response = RetryResponse(endpoint2, options);

  //リクエスト結果
  console.log(response)
}

const SheetNames = Object.freeze({
  SKIP: "skip",
  BOOKING_TARGET: "bookingTarget",
  BOOKING_TARGET_FOLLOWING: "bookingTargetFollowing",
  BOOKING_TARGET_FOLLOWERS: "bookingTargetFollowers",
  MY_FOLLOWING: "myFollowing",
  MY_FOLLOWERS: "myFollowers",
  ALPHA: "alpha",
  ACTIVE: "active",
  FOLLOWING: "following",
  UNFOLLOWING: "unfollowing",
  REPORT: "report",
});

const FollowType = Object.freeze({
  FOLLOWING: "following",
  FOLLOWERS: "followers",
});

//自分の情報取得
function getMe(){
  //リクエスト実行
  var response = RetryResponse("https://api.twitter.com/2/users/me"+"?user.fields=public_metrics"+"&expansions=pinned_tweet_id"); 
  //リクエスト結果
  var data = response["data"]
  console.log(data)
  return data
}

//自分のリスト情報取得
function getMelist(){
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/users/${ME_USER_ID }/owned_lists`); 
  //リクエスト結果
  console.log(response)
}

// 保護リストから保護フォロー一覧を取得
function getGuardFollowing()
{
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/lists/${ME_LIST_ID }/members`); 
  //リクエスト結果
  var ids = []
  for(user of response["data"])
  {
    ids.push(user["id"])
  }
  
  console.log(`保護フォロー一覧を${ids.length}名取得しました` )
  return ids
}

// 自分のTL取得
function getTimelines()
{
  //リクエスト実行
  var url = Utilities.formatString("https://api.twitter.com/2/users/%s/tweets",ME_USER_ID );
  var response = RetryResponse(url); //+"?max_results=50"
  var timelineTweetsData = response["data"]

  //リクエスト結果
  console.log(`自分の最新Tweetを${timelineTweetsData.length}件取得しました` )
  console.log(timelineTweetsData)
  return timelineTweetsData
}

// 自分のTLをいいねしたユーザー一覧
function getMeTweetlikeingUsers()
{
  var likeingUserdatas=[];
  //リクエスト実行
  // TLからいいねされたユーザーを取り出す
  for(imeline of getTimelines())
  {
    var url = Utilities.formatString("https://api.twitter.com/2/tweets/%s/liking_users",imeline["id"]);
    var response = RetryResponse(url);  
   // console.log(response["data"])

    //配列を連結する
    if(response["data"])
    {   
      likeingUserdatas.push(response["data"])
    }
  }

  var url = Utilities.formatString("https://api.twitter.com/2/tweets/%s/liking_users",ME_PINNED_TWEET_ID );
  var response = RetryResponse(url);  

  //配列を連結する
  if(response["data"])
  {   
    likeingUserdatas.push(response["data"])
  }

  // 特定の情報のみ配列にし直す
  var likeingUsersName=[];
  var likeingUsersId=[];
  var likeingUsers=[];
  for(likeingUserdata of likeingUserdatas)
  {
     likeingUsersName.push(likeingUserdata["name"])
     likeingUsersId.push(likeingUserdata["id"])
     likeingUsers.push([likeingUserdata["id"],likeingUserdata["name"]])
  } 
  
  //リクエスト結果
 // console.log(likeingUsersName.filter((x,i,array)=> array.indexOf(x) == i))
 //被りを削除
  likeingUsers = Array.from(new Set([...likeingUsers].map(JSON.stringify))).map(JSON.parse)
  console.log(`自分の最新Tweetにいいねしたユーザーを${likeingUsers.length}名取得しました` )
  return likeingUsers
 // return likeingUsers.filter((x,i,array)=> array.indexOf(x) == i)
}

function MeTweetlikeingUsersfollowing()
{
  var sheetIds = getSheetIds(SheetNames.FOLLOWING)
  var likeingUsers = getMeTweetlikeingUsers()
  var followingUsers = likeingUsers.filter(item => sheetIds.includes(item[0])==false)
  console.log(`新規でフォローするアカウントを${followingUsers.length}名取得しました` )
  if(followingUsers.length==0)
  {
    throw new Error("フォロー出来るアカウントがありません");
  }
  
  var count = 0
  for(user of followingUsers)
  {
    var userId=user[0]
    var name=user[1]

    console.log(`フォロー制限防止の為1名のみフォロー実行します` )
    Utilities.sleep(10000)//10秒待ってからフォローする
    count++
    // console.log(`全体進捗：${count}/${followingUsers.length}` )
    following(userId)
    setSheet(userId,SheetNames.FOLLOWING)
     console.log(`${name}をフォローしました` )
     return
  } 
}

function FindTweetsUsersfollowing()
{
  var sheetIds = getSheetIds(SheetNames.FOLLOWING)
  var userIds = findTweetsUserIds("#VRChat始めてました OR #VRChat始めました OR 自作トラッカー lang:ja",100)
  var followingUsers = userIds.filter(item => sheetIds.includes(item)==false)
  console.log(`新規でフォローするアカウントを${followingUsers.length}名取得しました` )
  if(followingUsers.length==0)
  {
    throw new Error("フォロー出来るアカウントがありません");
  }

  for(userId of followingUsers)
  {
    var rate = getFFRate(userId)
    if( rate > 1.1 )
    {
      console.log(`FF比が1.1以上の為飛ばします` )
      continue
    }

    console.log(`フォロー制限防止の為1名のみフォロー実行します` )
    Utilities.sleep(10000)//10秒待ってからフォローする
    following(userId)
    setSheet(userId,SheetNames.FOLLOWING)
    console.log(`ユーザーID${userId}をフォローしました` )
    return
  }
}

// エラーユーザーIDをシートから削除する
function DeleteSheetDeathUsers()
{
  var sheetIds = getSheetIds(SheetNames.FOLLOWING)
  var errorIds = getErrorUsersData(sheetIds)

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SheetNames.FOLLOWING);

  for(errorId of errorIds)
  {
    sheetIds = getSheetIds(SheetNames.FOLLOWING)
    var deleteIndex = sheetIds.indexOf(errorId)
    if(0>deleteIndex)
    {
      continue
    }
    sheet.deleteRow(deleteIndex+1)
    
    console.log(`${errorId}はエラーIDの為、${deleteIndex+1}行目を削除しました。` )
  }
}

//非アクティブユーザーのフォローを外す
function InactiveUserUnfollow()
{
  var guardFollowingIds = getGuardFollowing()
  var activeIds = getSheetIds(SheetNames.ACTIVE)
  var followingIds = getFollowList(ME_USER_ID  , FollowType.FOLLOWING); 
  var deleteFollowingIds = followingIds.filter(item=>guardFollowingIds.includes(item)==false&&activeIds.includes(item)==false)
  deleteFollowingIds = deleteFollowingIds.reverse()
  
  if(deleteFollowingIds.length==0)
  {
    throw new Error("非アクティブアカウントがありませんでした");
  }

  console.log(`アクティブ未判定アカウントを${deleteFollowingIds.length}名取得しました。` ) 

  for(userId of deleteFollowingIds)
  {
    var newTweetDate = getNewTweetDate(userId)

    // 1ヶ月前の日時取得
    var date = getOneMonthAgoDatetime()

    if(newTweetDate==null || date<newTweetDate)
    {
      console.log(`最終ツイートが1ヶ月以内でアクティブの為飛ばします` )
      setSheet(userId,SheetNames.ACTIVE)
      continue
    }

    var likedNewTweetDate = getLikedNewTweetDate(userId)
    if( likedNewTweetDate == null || date<likedNewTweetDate )
    {
      console.log(`最新いいねしたツイートが1ヶ月以内でアクティブの為飛ばします` )
      setSheet(userId,SheetNames.ACTIVE)
      continue
    }

    console.log(`最終ツイートが1ヶ月以上前で非アクティブの為フォローを外します。` )
    console.log(`制限防止の為1名のみ実行します` )
    Utilities.sleep(10000)//10秒待つ
    deleteFollowing(userId)
    console.log(`ユーザーID${userId}をリムーブしました` )
    setSheet( userId, SheetNames.UNFOLLOWING )
    return
  }
}

// フォロー中が多いアカウントのフォローを外す
function SpamUserUnfollow()
{ 
  var alphaIds = getSheetIds(SheetNames.ALPHA)

  // 1週間前の日時取得
  var date = getOneWeekAgoDatetime()

  //保護リスト
  var guardFollowingIds = getGuardFollowing()
  
  //フォロー中
  var followingIds = getFollowList(ME_USER_ID  , FollowType.FOLLOWING); 
  
  //フォローしたユーザー日時データ
  var sheetDatas = getSheetDatas(SheetNames.FOLLOWING)

  //フォロー解除リスト
  var list = sheetDatas.filter(item=>guardFollowingIds.includes(item[0])==false&&followingIds.includes(item[0])&&item[1]<date&&alphaIds.includes(item[0])==false)
  
  console.log( `フォロー解除リストを${list.length}名取得`)
  for( userData of list )
  { 
    var rate = getFFRate(userData[0])
    var followersNum = getFollowersNum(userData[0])
    if( rate < 0.9 || followersNum < 1000)
    {
      Utilities.sleep(10000)//10秒待つ
      deleteFollowing(userData[0])
      console.log(`ユーザーID【${userData[0]}】をリムーブしました` )
      setSheet( userData[0], SheetNames.UNFOLLOWING )
      return
    }

    console.log(`FF比が0.9以上&フォロワーが1000以上でした` )
    setSheet(userData[0],SheetNames.ALPHA)
  }
}

//片思いユーザーのフォローを 外す
function KataomoiUserUnfollow()
{ 
  // 1週間前の日時取得
  var date = getOneWeekAgoDatetime()

  //保護リスト
  var guardFollowingIds = getGuardFollowing()
  
  //片思いリスト
  var followersIds = getSheetIds(SheetNames.MY_FOLLOWERS)
  var followingIds = getSheetIds(SheetNames.MY_FOLLOWING)
  followersIds.shift()
  followingIds.shift()
  var kataomoiIds = followingIds.filter( item => followersIds.includes( item ) == false )
  
  //フォローしたユーザー日時データ
  var sheetDatas = getSheetDatas(SheetNames.FOLLOWING)

  //フォロー解除したユーザーリスト
  var unfollowing = getSheetIds(SheetNames.UNFOLLOWING)

  //フォロー解除リスト
  var list = sheetDatas.filter(item=>unfollowing.includes(item[0])==false&&guardFollowingIds.includes(item[0])==false&&kataomoiIds.includes(item[0])&&item[1]<date)
  
  console.log( `フォロー解除リストを${list.length}名取得`)
  for( userData of list )
  {
    console.log(`制限防止の為1名のみ実行します` )
    Utilities.sleep(10000)//10秒待つ
    deleteFollowing(userData[0])
    console.log(`ユーザーID【${userData[0]}】をリムーブしました` )
    setSheet( userData[0], SheetNames.UNFOLLOWING )
    return
  }
}

function FindFriendOfFriendFollowing()
{
  var sheetIds = getSheetIds(SheetNames.FOLLOWING)
  var skipIds = getSheetIds(SheetNames.SKIP)
  
  var followersIds = getSheetIds(SheetNames.BOOKING_TARGET_FOLLOWERS)
  var followingIds = getSheetIds(SheetNames.BOOKING_TARGET_FOLLOWING)
  followersIds.shift()
  followingIds.shift()
  var userIds = followersIds.filter(item => followingIds.includes(item))
  
  var followingUsers = userIds.filter(item => sheetIds.includes(item)==false&&skipIds.includes(item)==false)
  console.log(`新規でフォローするアカウントを${followingUsers.length}名取得しました` )
  if(followingUsers.length==0)
  {
    deleteRowAt(SheetNames.BOOKING_TARGET)
    throw new Error("フォロー出来るアカウントがありません");
  }

  for(userId of followingUsers)
  {
    var rate = getFFRate(userId)
    if( rate > 1.1 )
    {
      console.log(`FF比が1.1以上の為飛ばします` )
      setSheet(userId,SheetNames.SKIP)
      continue
    }

    var newTweetDate = getNewTweetDate(userId)

    // 24時間前の日時取得
    var date = getOneDayAgoDatetime()

    if( newTweetDate != null && date < newTweetDate )
    {
      followingAndSetSheet(userId)
      return
    }

    var likedNewTweetDate = getLikedNewTweetDate(userId)
    if( likedNewTweetDate != null && date < likedNewTweetDate )
    {
      followingAndSetSheet(userId)
      return
    }
    
    console.log(`最終Twitter更新が24時間以上前の為飛ばします` )
    setSheet(userId,SheetNames.SKIP)
  }
}

// ユーザーをフォローしてフォローシートに記載する
function followingAndSetSheet( userId )
{
  Utilities.sleep(10000)//10秒待ってからフォローする
  following( userId )
  setSheet( userId, SheetNames.FOLLOWING )
  console.log(`ユーザーID${userId}をフォローしました` )
}

function ReportFF()
{
  console.log(`本日のFF数を保存します` )
  var me = getMe()
  setSheetFF(me)
}

// FF数をスプレットシートに書き込む
function setSheetFF(meData)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SheetNames.REPORT);
  var meMetrics = meData["public_metrics"]
  var followingCount = meMetrics["following_count"]
  var followersCount = meMetrics["followers_count"]
  var totalFollowingCount = getSheetIds(SheetNames.FOLLOWING)
  sheet.appendRow([new Date(),followingCount,followersCount,totalFollowingCount.length]);
  console.log(`FF数を保存しました` )
  console.log(`フォロー数：${followingCount}フォロワー数：${followersCount}歴代フォロー数：${totalFollowingCount.length}` )
}

function manualFindTweetsUsersfollowing()
{
  var sheetIds =getSheetIds(SheetNames.FOLLOWING)
  var userIds= findTweetsUserIds("#VRC個人開発集会",10)
  for(userId of userIds)
  {
    // 既にフォロー済みであれば何もしない
    if(sheetIds.includes(userId))
    {
      continue
    }

    Utilities.sleep(10000)//10秒待ってからフォローする
    following(userId)
    console.log(`ユーザーID${userId}をフォローしました` )
    setSheet(userId,SheetNames.FOLLOWING)
  }
}

// FFから特定ユーザーのリストを返す
function getUserList( id, status )
{
  //リクエスト実行
  var followingList = getFollowList(id , FollowType.FOLLOWING); 
  var followersList = getFollowList(id , FollowType.FOLLOWERS); 

  var ids = []
  switch(status)
  {
    case "ryoomoi":
        ids = followingList.filter( item => followersList.includes( item ) )
      break;

    case "kataomoi":
        ids = followingList.filter( item => followersList.includes( item ) == false )
      break;

    case "kataomoware":
        ids = followersList.filter( item => followingList.includes( item ) == false )
      break;
  }

  //リクエスト結果
  console.log(`【${status}】は${ids.length}名いました。`)
  return ids
}

function getFollowList(id, status) 
{
  var ids = [];
  var nextToken = "";
  var url = `https://api.twitter.com/2/users/${id}/${status}`;

  do 
  {
    // リクエスト実行
    var requestUrl = url + "?max_results=1000&user.fields=protected";
    if (nextToken !== "") 
    {
      requestUrl += `&pagination_token=${nextToken}`;
    }

    var response = RetryResponse(requestUrl);

    // リクエスト結果
    for (const user of response["data"]) 
    {
      if (user["protected"]) 
      {
        continue;
      }
      ids.push(user["id"]);
    }

    nextToken = response["meta"]["next_token"];
  } 
  while (response["meta"]["next_token"]);

  console.log(`公開アカウントの【${status}】は${ids.length}名いました。`);
  return ids;
}

function getErrorUsersData(ids)
{
  var errorIds = []
 
  while(0<ids.length)
  {
    var chopIds = ids.splice(0,100)
    //リクエスト実行
    var response = RetryResponse(`https://api.twitter.com/2/users?ids=${chopIds.join(`,`)}`); 

    if(response["errors"])
    {
      for(data of response["errors"])
      {
        errorIds.push(data["value"])
      }
    }
  }
  //リクエスト結果
  console.log(`エラーIDが${errorIds.length}名いました。`)
  return errorIds
}

// 最新ツイートの日時取得
function getNewTweetDate(id)
{
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/users/${id}/tweets?tweet.fields=created_at`); 

  if(response["data"])
  {
    var data = response["data"]
    var date = data[0]["created_at"]

    //リクエスト結果
    console.log(`【${id}】の最新ツイートの日時は${date}`)
    return new Date(date)
  }
  
  console.log(`【${id}】の最新ツイートは存在しませんでした。`)
  return null
}

// ユーザーが気に入ったツイートの最新日時（最新100件内）
function getLikedNewTweetDate( id )
{
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/users/${id}/liked_tweets?tweet.fields=created_at`); 

  if( response["data"] )
  {
    var datas = response["data"]
    console.log(`【${id}】のいいねしたツイートの最新${datas.length}件を取得しました。`)
    var newDate
    
    for( data of datas )
    {
      var date = new Date(data["created_at"])
      if( newDate == null ||　newDate < date)
      {
        newDate = date
      }
    }

    //リクエスト結果
    console.log(`【${id}】がいいねしたツイートの最新日時は${newDate}`)
    return newDate
  }
  
  console.log(`【${id}】がいいねしたツイートは存在しませんでした。`)
  return null
}

function getFFRate(id)
{
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/users/${id}`+"?user.fields=public_metrics"); 

  var data = response["data"]
  var name = data["name"]
  var metrics = data["public_metrics"]
  var followers = metrics["followers_count"]
  var following = metrics["following_count"]
  //リクエスト結果
  var rate = followers / following
  console.log(`名前：${name}　FF比：${rate}`)
  return rate
}

function getFollowersNum(id)
{
  //リクエスト実行
  var response = RetryResponse(`https://api.twitter.com/2/users/${id}`+"?user.fields=public_metrics"); 

  var data = response["data"]
  var name = data["name"]
  var metrics = data["public_metrics"]
  var followers = metrics["followers_count"]
  console.log(`名前：${name}　フォロワー数：${followers}`)
  return followers
}

// ユーザーIDをスプレットシートに書き込む
function setSheet(userId,sheetName)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  sheet.appendRow([userId.toString(),new Date()]);
}

// スプレットシートからユーザーIDリスト読み込む
function getSheetIds(sheetName)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  if(1>lastRow)
  {
    console.log(`【${sheetName}】IDリストを0件取得しました` )
    return []
  }
  var ranges = sheet.getRange(1, 1, lastRow).getValues();

  //2次元配列なので1次元配列にする
  var ids = []
  for(range of ranges)
  {
    ids.push(range.toString())
  } 
  
  console.log(`【${sheetName}】IDリストを${ids.length}件取得しました` )
  return ids
}

//スプレットシートからユーザーリスト読み込む
function getSheetDatas(sheetName)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var ranges = sheet.getRange(1, 1, lastRow,2).getValues();
  
  //型がずれている場合があるので揃える
  var datas = []
  for(range of ranges)
  {
    datas.push([range[0].toString(),new Date(range[1])])
  } 
  
  console.log(`【${sheetName}】Dataリストを${datas.length}件取得しました` )
  return datas
}

// フォローリストを取得してシートに記載
// 手動フォローはリストに記載されない為
function SetSheetUsersfollowing()
{
  var sheetIds = getSheetIds(SheetNames.FOLLOWING)
  var followingList = getFollowList(ME_USER_ID  , FollowType.FOLLOWING); 
  var ids = followingList.filter(item=>sheetIds.includes(item)==false)
    console.log(`${ids.length}名未記載のユーザーを発見しました` )
  for(id of ids)
  {
    setSheet(id,SheetNames.FOLLOWING)
    console.log(`ユーザーID${id}をフォローシートに記載しました` )
  } 
}

function getFollowings()
{
  var ids = []
  var nextToken=""
  var url = Utilities.formatString("https://api.twitter.com/2/users/%s/following",ME_USER_ID );
  do
  {
    //リクエスト実行
    var requestUrl = url+"?max_results=1000"
    if(nextToken!="")
    {
      requestUrl +=`&pagination_token=${nextToken}`
    }
    var response = RetryResponse(requestUrl); 
    //リクエスト結果
    for(user of response["data"])
    {
      ids.push(user["id"])
    } 
    console.log(response)
    nextToken=response["meta"]["next_token"]
  }
  while(response["meta"]["next_token"])
  
  console.log(ids)
  return ids
}

function findTweetsUserIds(word,num)
{
  // API URL
  var getPoint = 'https://api.twitter.com/2/tweets/search/recent?'
  // 検索キーワードとパラメータ
  var keyWord = '&query='+ encodeURIComponent(word)
  var params = 'expansions=author_id'+`&max_results=${num}`
  
  var result = RetryResponse(getPoint+params+keyWord)
  
  var ids=[]
  for(tweet of result["data"])
  {
     ids.push(tweet["author_id"])
  }

  ids = ids.filter((x,i,array)=> array.indexOf(x) == i)
  console.log(`検索ワード「${word}」の最新${num}件投稿したユーザーを${ids.length}名取得しました`)
  return ids
}


function following(targetId){
  var following_ep = `https://api.twitter.com/2/users/${ME_USER_ID }/following`;

   var data = {
    "target_user_id" : targetId
  }

  var options = {
    "method": "post",
    'contentType': 'application/json',
    'accept': "application/json",
    'payload': JSON.stringify(data)
  }

  //リクエスト実行
  var response = RetryResponse(following_ep,options);

  console.log(response)
}

function deleteFollowing(targetId){
  var following_ep = `https://api.twitter.com/2/users/${ME_USER_ID }/following/${targetId}`;

  var options = {
    "method": "DELETE",
    'contentType': 'application/json',
    'accept': "application/json"
  }
  //リクエスト実行
  var response = RetryResponse(following_ep,options);

  console.log(response)
}

function RetryResponse(url,options=null)
{
  //トークン確認
  var service = checkOAuth(appname);
  var count=0
  while(count<5)
  {
    count++
    try
    {
      //リクエスト実行
      return JSON.parse(service.fetch(url,options)); 
    }
    catch(error)
    {
      if(error.message.includes("Address unavailable")==false||count>=5)
      {
        console.log(url)
        throw new Error(error.message);
      }
      
      console.log(error.message)
      Utilities.sleep(10000)//10秒待ってから再実行
    }
  }
}

function mainFollowing()
{
  var targetFollowSearch = (sheetNames,followType)=>
  {
    var targetId 
    var nextToken

    var bookingData = getRowAt(SheetNames.BOOKING_TARGET)
    var targetData = getRowAt(sheetNames)

    if(targetData.targetId==bookingData.targetId)
    {
      targetId=targetData.targetId;
      nextToken=targetData.nextToken;
      if(isNullOrEmpty(nextToken))
      {
        console.log(sheetNames+"シートは取得済みです")
        return
      }
    }
    else
    {
      sheetClear(sheetNames)
      targetId=bookingData.targetId;
      nextToken=bookingData.nextToken;
    }

    do
    {
      var followers = getFollowListNextToken(followType,targetId,nextToken)
      updateRowAt(sheetNames,targetId,followers.nextToken)
      addIds(sheetNames,followers.ids)
      nextToken=followers.nextToken
    }
    while(isNullOrEmpty(nextToken)==false)
  }
  targetFollowSearch(SheetNames.BOOKING_TARGET_FOLLOWERS,FollowType.FOLLOWERS)
  targetFollowSearch(SheetNames.BOOKING_TARGET_FOLLOWING,FollowType.FOLLOWING)

  FindFriendOfFriendFollowing()
}

function mainUnfollowing()
{
  var myFollowSearch = (sheetNames,followType)=>
  {
    var nextToken
    var followData = getRowAt(sheetNames)

    if(isNullOrEmpty(followData.targetId) || followData.date<getOneDayAgoDatetime())
    {
      console.log(sheetNames+"シートを更新します")
      sheetClear(sheetNames)
    }
    else if(isNullOrEmpty(followData.nextToken)==false)
    {
      console.log(sheetNames+"シートに追加します")
      nextToken=followData.nextToken;
    }
    else
    {
      console.log(sheetNames+"シートは取得済みです")
      return
    }

    do
    {
      var followers = getFollowListNextToken(followType,ME_USER_ID,nextToken)
      updateRowAt(sheetNames,ME_USER_ID,followers.nextToken)
      addIds(sheetNames,followers.ids)
      nextToken=followers.nextToken
    }
    while(isNullOrEmpty(nextToken)==false)
  }
  myFollowSearch(SheetNames.MY_FOLLOWERS,FollowType.FOLLOWERS)
  myFollowSearch(SheetNames.MY_FOLLOWING,FollowType.FOLLOWING)
  KataomoiUserUnfollow()
}

function sheetClear(sheetName)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  sheet.clear()
}

function isNullOrEmpty(string) 
{
  return string == null || string == "";
}

function getOneWeekAgoDatetime() 
{
  // 1週間前の日時取得
  var date = new Date();
  return date.setDate(date.getDate() - 7);
}

function getOneDayAgoDatetime() {
  // 1日前の日時取得
  var date = new Date();
  return new Date(date.setDate(date.getDate() - 1));
}

function getOneMonthAgoDatetime() {
  // 1ヶ月前の日時取得
  var date = new Date();
  return new Date(date.setMonth(date.getMonth() - 1));
}

// シートにターゲットユーザーとnextTokenと更新日付を記載
function updateRowAt(sheetName,targetId,nextToken) 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var rowData = [targetId,nextToken,new Date()]
  var range = sheet.getRange(1, 1, 1, rowData.length);

  // データを上書き
  range.setValues([rowData]);
  console.log(`【${sheetName}】シートにid【${targetId}】とnextToken【${nextToken}】を記載` )
}

// シートにidsの追加
function addIds(sheetName,ids)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var range = sheet.getRange(sheet.getLastRow()+1, 1, ids.length);
  var idData = []
  for(id of ids)
  {
    idData.push([id])
  }
  range.setValues(idData);
  console.log(`【${sheetName}】シートにidを【${ids.length}】件記載` )
}

function deleteRowAt(sheetName)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  sheet.deleteRow(1)
}

// シートからターゲットユーザーとnextTokenを取得
function getRowAt(sheetName) 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var data = sheet.getRange(1, 1, 1, 3).getValues();
  return {"targetId":data[0][0],"nextToken":data[0][1],"date":new Date(data[0][2])}
}

// ターゲットユーザーのフォローリストとnextTokenを取得
function getFollowListNextToken(status,targetId,nextToken) 
{
  var ids = [];
  var url = `https://api.twitter.com/2/users/${targetId}/${status}`;

  // リクエスト実行
  var requestUrl = url + "?max_results=1000&user.fields=protected";
  if (isNullOrEmpty(nextToken)==false) 
  {
    requestUrl += `&pagination_token=${nextToken}`;
  }
  
  var response = RetryResponse(requestUrl);

  // リクエスト結果
  for (const user of response["data"]) 
  {
    if (user["protected"]) 
    {
      continue;
    }
    ids.push(user["id"]);
  }

  nextToken = response["meta"]["next_token"];
  
  console.log(`【${targetId}】の【${status}】リストを【${ids.length}】件取得しました` )
  return {"ids":ids,"nextToken":nextToken}
}

//ログアウト
function reset() {
  OAuth1.createService(appname)
      .setPropertyStore(PropertiesService.getUserProperties())
      .reset();
  SpreadsheetApp.getUi().alert("ログアウトしました。")
}
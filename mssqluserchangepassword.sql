GO
	DECLARE @user_id nvarchar(255), @user_name nvarchar(255), @last_login_date smalldatetime,
		@email_address varchar(255), @bad_count nvarchar(10), @disabled int, @distinguishedName varchar(255);
	DECLARE @body nvarchar(MAX);
	DECLARE @no_mail int;
DECLARE users CURSOR LOCAL FAST_FORWARD FOR
SELECT sAMAccountName
, displayName
, DATEADD(minute, DATEDIFF(minute, GETUTCDATE(), GETDATE()), CASE WHEN CAST([pwdLastSet] AS bigint) > 0 THEN CAST([pwdLastSet] AS bigint) / (864000000000.0) ELSE 109207 END)
		 - 109207 AS pwdLastSet
, ISNULL(mail,'') AS mail
, LEFT(CAST(SUBSTRING(CONVERT(binary(10), [userAccountControl]),8,1)+0 AS char ),1) AS neverExpire
, userAccountControl & 2 AS Disabled 
, distinguishedName
FROM OpenQuery
(
ADSI, 
'SELECT pwdLastSet, userAccountControl, displayName, sAMAccountName,
mail, distinguishedName
FROM  ''LDAP://subdomain.domain.corp/OU=Users,OU=OUUsers,DC=subdomain,DC=domain,DC=CORP''
WHERE objectClass =  ''User'' AND objectCategory = ''Person'' 
'
) AS tblADSI1 WHERE userAccountControl & 2 = 0 AND SUBSTRING(CONVERT(binary(10), [userAccountControl]),8,1)+0 = 0 AND LEN(mail) > 0 AND DATEADD(minute, DATEDIFF(minute, GETUTCDATE(), GETDATE()), CASE WHEN CAST([pwdLastSet] AS bigint) > 0 THEN CAST([pwdLastSet] AS bigint) / (864000000000.0) ELSE 109207 END)
		 - 109207 < DATETIMEFROMPARTS (YEAR(GETDATE()), MONTH(GETDATE()), DAY(GETDATE()), 23, 59, 59, 0) -29

	OPEN users
	FETCH NEXT FROM users INTO @user_id, @user_name, @last_login_date, @email_address, @bad_count, @disabled, @distinguishedName
WHILE @@FETCH_STATUS = 0

BEGIN
	IF DATEDIFF(dd,DATETIMEFROMPARTS (YEAR(GETDATE()), MONTH(GETDATE()), DAY(GETDATE()), 23, 59, 59, 0), @last_login_date + 93 - 2) IN (14,7,6,5,4,3,2,1)
		BEGIN

		IF LEN(@email_address) = 0 OR @user_id IN ('jirasupport')
			BEGIN
				SET @email_address = @email_address
			END
		PRINT N'Отправка письма для ' + @email_address + N' for ' + @user_name + ' ...'

		SET @body = N'
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<title>Действие пароля скоро истекает</title>
</head>
<body>
<p><img style="float:right;" src="https://domain.com/logo.png"/></p>
<p>Здравствуйте, ' + @user_name + N'!</p>
<p>Действие пароля для учетной записи ' + @user_id + ' истекает
 <b>' + CONVERT(varchar,@last_login_date + 93 - 2,120) + '</b>. Через <b> ' + CONVERT(varchar,DATEDIFF(dd,DATETIMEFROMPARTS (YEAR(GETDATE()), MONTH(GETDATE()), DAY(GETDATE()), 23, 59, 59, 0), @last_login_date + 93 - 2)) + ' </b> дней вход будет невозможен, а соответственно и сменить пароль самостоятельно не удастся.
Поэтому его необходимо заменить до этого числа. Для смены пароля воспользуйтесь одним из следующий способов.</p>
<p><b>Первый способ.</b> Войдите в https://owa.domain.com/owa/auth/logon.aspx?replaceCurrent=1&url=https%3A%2F%2Fowa.domain.com%2Fowa%2F%23path%3D%2Foptions%2Fmyaccount под учетной записью <b>'+ SUBSTRING(@distinguishedName, 
        CHARINDEX('DC=', @distinguishedName)+3,
        CHARINDEX('DC=', SUBSTRING(@distinguishedName,CHARINDEX('DC=', @distinguishedName)+5,LEN(@distinguishedName)))
 )+'\'+ @user_id +'</b> и нажмите "Изменить пароль".<br>
Введите текущий пароль и 2 раза новый.</p>
<p><b>Второй способ.</b> Если вы работает на рабочей станции в домене, то нажмите <b>CTRL+ALT+DELETE</b> и нажмите "Изменить пароль". Если вы работаете через службу удаленных рабочих столов, то сочетание клавиш <b>CTRL+ALT+END</b>.<br>
Введите текущий пароль и 2 раза новый.</p>
<p>Если учетная запись заблокирована, то обратитесь к администраторам</p>
<p><b>Пароль должен удовлетворять следующим минимальным требованиям.</b></p>
<p>Не содержать имени учетной записи пользователя или частей полного имени пользователя длиной более двух рядом стоящих знаков<br>
Иметь длину не менее 12 знаков<br>
Содержать знаки трех из четырех перечисленных ниже категорий:<br>
Латинские заглавные буквы (от A до Z)<br>
Латинские строчные буквы (от a до z)<br>
Цифры (от 0 до 9)<br>
Отличающиеся от букв и цифр знаки (например, !, $, #, %)<br>
Пароль не должен совпадать с одним из 24-х ранее использованнх.</p>
</body>
</html>';

		EXEC msdb.dbo.sp_send_dbmail
			@recipients = @email_address,
			@subject = N'Действие пароля скоро истекает',
			@body = @body,
			@body_format = 'HTML',
			@profile_name = 'user_not';

		WAITFOR DELAY '00:00:03';
		END
		FETCH NEXT FROM users INTO @user_id, @user_name, @last_login_date, @email_address, @bad_count, @disabled, @distinguishedName
END
CLOSE users
GO

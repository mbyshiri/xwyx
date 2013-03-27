<%
'如果网站频道启用子域名功能，则需要修改下面的设置
Const Enable_SubDomain = False          '子域名功能开关 True=启用，False=禁用
Const DomainRoot = "powereasy.net"      '网站域名根
Const strSubDomains = "www|news|shop|soft"   '主机名（子域名）列表。比如要启用"news.powereasy.net"，这里就加上"news"，多个子域名之间用半角"|"分隔

Const EnableStopInjection = True        '是否启用防SQL注入功能，True为启用，False为禁用
Const ShowUnpass = True           '后台待审核详细资料开关 True=启用，False=禁用
Const FriendSiteCheckCode = True  '友情连接验证码开关 True=启用，False=禁用

Const EnableSiteManageCode = True       '是否启用后台管理认证码 是： True  否： False
Const SiteManageCode = "PowerEasy2008"  '后台管理认证码，您可以修改成您的管理员认证码：×××××××××

Const MaxPerPage_Create = 10   '生成HTML时，每页生成的数量，建议不要超过100，否则可能会导致页面超时
Const SleepTime = 3            '每页生成完毕后，暂停时间，单位为秒。如果为0，则不暂停，生成当前页面后马上跳转到下一页继续生成。建议设置为3－10
%>

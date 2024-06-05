--declare @year int; set @year = 2022
--declare @quarter int; set @quarter = 1
--declare @CarrierID varchar(3); set @carrierID = 'JHL'

declare @month int
set @month = @quarter*3
declare @YearMonth varchar(6)
set @yearmonth = concat(@year, right(concat('0',@month),2))



declare @ID varchar(3)
set @ID = '1JH'

select @carrierid CarrierID
,@quarter [Quarter], ReinsurerCompany,format(sum(totalPremium),'N2') 'Total Premium'
	from vbilling

where cedingcompany = @ID
and dateBilled IN (@YearMonth,@YearMonth-1,@YearMonth-2)
and cedingcompany+reinsurancetreaty in (
'PRUSQMAGGIB','1JHSWMAGF+A','1JHSWMAGS+C','1JHINMAGS+C','1JHLNMAGS+B','1JHSWMAGS+B','1JHSWMAGS-C',
'1JHINMAGS-C','1JHSWMAGF-A','1JHINMAGF-A','1JHINMAGF+A','1JHSWMBGUNE','1JHSWMBGS-E','1JHSWMBGS+E',
'1JHSWMAGF+B','1JHSWMBGS+F','1JHLQMAGGIA','1JHRQMAGGIA','1JHLNMAGGIA','1JHRGMAGGIA','1JHSWMAGFUA',
'1JHRGMAGFUA','1JHLNMAGUNA','1JHRGMAGUNA','1JHLNMAGUNB','1JHSWMAGUNB','1JHTAMAGFUA','1JHLQMAGGIB',
'1JHSQMAGGIB','1JHLNMAGGIB','1JHSWMAGGIB','1JHSWMAGUNC','1JHINMAGUNC','1JHINMAGFUA','1JHSQMAGGIC',
'1JHIQMAGGIC','1JHSWMAGGIC','1JHINMAGGIC','1JHSQMBGGIE','1JHSWMBGGIE','1JHSQMAGGIE','1JHCQMAGGIA',
'1JHRQMAGGIC','1JHSWMAGGIE','1JHCTMAGGIA','1JHRGMAGGIC','1JHSWMBGUNF','1JHSWMAGFUB','1NYINMAGGIC',
'1NYINMAGUNC','1NYIQMAGGIC','1NYMQMAGGKC','1NYMUMAGGKC','1NYMUMAGGKD','1NYSQMAGGIC','1NYSQMAGGKC',
'1NYSWMAGGIC','1NYSWMAGGKC','1NYSWMAGGKD','1NYSWMAGUNC','1PMSWMAGUNE','1PMRGMAGUNC','1PMSWMAGUNC',
'1PMINMAGUNC','1PMSWMAGS+D','1PMRGMAGS+B','1SLSLMAGFUA','1SLSLMAGF+A','PRUSWMAGGIB','PRUSWMAGFUA',
'PRUINMAGFUA','PRUSWMAGS+C','PRUINMAGS+C','PRUSWMAGS-E','PRURGMAGS-C','PRUSWMAGS+E','PRURGMAGS+C',
'PRURGMAGFUB','PRUSWMAGFUB','PRUSWMAGUNE','PRURGMAGUNC','1SLSLMAGF+B','1SLSLMAGF-A','1SLSLMAGF-B',
'1SLSLMAGFUB','1SLSLMAGTPA','1SLSLFACMAG','1SLSLFACTPA','1SLSLFACTPB')

group by reinsurercompany
order by reinsurercompany

--YRT Detail
select b.policyNumber, format(sum(b.totalPremium),'N2'), i.[state], b.ReinsurerCompany
	from vbilling b
	left join (
		select distinct policynumber, [state] 
		from inforce 
		where cedingCompany = '1JH'
		and plancode like '%MAG%') i
	on b.policynumber = i.policynumber

where cedingcompany = @ID
and dateBilled IN (@YearMonth,@YearMonth-1,@YearMonth-2)
and cedingcompany+reinsurancetreaty in (
'PRUSQMAGGIB','1JHSWMAGF+A','1JHSWMAGS+C','1JHINMAGS+C','1JHLNMAGS+B','1JHSWMAGS+B','1JHSWMAGS-C',
'1JHINMAGS-C','1JHSWMAGF-A','1JHINMAGF-A','1JHINMAGF+A','1JHSWMBGUNE','1JHSWMBGS-E','1JHSWMBGS+E',
'1JHSWMAGF+B','1JHSWMBGS+F','1JHLQMAGGIA','1JHRQMAGGIA','1JHLNMAGGIA','1JHRGMAGGIA','1JHSWMAGFUA',
'1JHRGMAGFUA','1JHLNMAGUNA','1JHRGMAGUNA','1JHLNMAGUNB','1JHSWMAGUNB','1JHTAMAGFUA','1JHLQMAGGIB',
'1JHSQMAGGIB','1JHLNMAGGIB','1JHSWMAGGIB','1JHSWMAGUNC','1JHINMAGUNC','1JHINMAGFUA','1JHSQMAGGIC',
'1JHIQMAGGIC','1JHSWMAGGIC','1JHINMAGGIC','1JHSQMBGGIE','1JHSWMBGGIE','1JHSQMAGGIE','1JHCQMAGGIA',
'1JHRQMAGGIC','1JHSWMAGGIE','1JHCTMAGGIA','1JHRGMAGGIC','1JHSWMBGUNF','1JHSWMAGFUB','1NYINMAGGIC',
'1NYINMAGUNC','1NYIQMAGGIC','1NYMQMAGGKC','1NYMUMAGGKC','1NYMUMAGGKD','1NYSQMAGGIC','1NYSQMAGGKC',
'1NYSWMAGGIC','1NYSWMAGGKC','1NYSWMAGGKD','1NYSWMAGUNC','1PMSWMAGUNE','1PMRGMAGUNC','1PMSWMAGUNC',
'1PMINMAGUNC','1PMSWMAGS+D','1PMRGMAGS+B','1SLSLMAGFUA','1SLSLMAGF+A','PRUSWMAGGIB','PRUSWMAGFUA',
'PRUINMAGFUA','PRUSWMAGS+C','PRUINMAGS+C','PRUSWMAGS-E','PRURGMAGS-C','PRUSWMAGS+E','PRURGMAGS+C',
'PRURGMAGFUB','PRUSWMAGFUB','PRUSWMAGUNE','PRURGMAGUNC','1SLSLMAGF+B','1SLSLMAGF-A','1SLSLMAGF-B',
'1SLSLMAGFUB','1SLSLMAGTPA','1SLSLFACMAG','1SLSLFACTPA','1SLSLFACTPB')

group by b.policyNumber, i.[state], ReinsurerCompany
order by reinsurercompany
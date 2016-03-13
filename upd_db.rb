require 'db_com'

def upd db, arr
	xh=arr[0]
	nf=arr[2].to_i
	wd=arr[3].to_i
	djpj=arr[5].to_i
	jsxh=arr[6].to_i
	pjr=arr[7]
	#~ p xh, nf, wd, djpj, jsxh, pjr
	sql="update JKSUZHI set djpj=#{djpj} where xh='#{xh}' and nf=#{nf} and wdxh=#{wd} 
	and jsxh=#{jsxh} and pjr='#{pjr}'"
	db.exec sql 
end

db2 = DB2.new({'host'=> "localhost\\sqlexpress", 'database'=> 'xx', 'username'=> 'sa', 'password'=>'sa'})
#~ db2 = DB2.new({'host'=> "TJ82SZPJ", 'database'=> 'xx', 'username'=> 'sa', 'password'=>'sasasasa'})
File.open('errorBook1.csv').each{|one|
	arr = one.split(/,/)
	if arr[10]=~/x/i
		upd db2, arr
	end 
}
db2.close

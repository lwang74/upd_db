require 'db_com'
require 'excel'
require 'cfg'

class CAll
	def initialize persion_like
		@persion_like=persion_like
	end
	
	def get_stds db
		hsh={}
		sql="select *
			from dbo.JKXUESHENG"
		db.exec_sql(sql){|row|
			hsh[row[0]]= row[2]
		}
		print "."
		hsh
	end

	def get_tch db
		hsh={}
		sql="select *
			from dbo.JKJIAOSHI"
		db.exec_sql(sql){|row|
			#~ p row
			hsh[row[0]]= row[2]
		}
		print "."
		hsh
	end

	def get_wd db
		hsh={}
		sql="select *
			from dbo.JKWEIDU"
		db.exec_sql(sql){|row|
			#~ p row
			hsh[row[0]]= row[1]
		}
		print "."
		hsh
	end

	def get_pj db
		hsh={}
		sql="select *
			from dbo.JKSUZHI sz
			where sz.xh like '#{@persion_like}%'"
		db.exec_sql(sql){|row|
			#~ p row
			hsh[row[0]] ||={}
			hsh[row[0]][row[2]] ||={}
			hsh[row[0]][row[2]][row[1]] ||={}
			hsh[row[0]][row[2]][row[1]][row[5]] ||=row[4]
		}
		print "."
		hsh
	end

	def get_bj db
		hsh={}
		sql="SELECT xh,bj.*
			FROM [dbo].[JKBAN_XUESHENG] bx
			join dbo.JKBANJI bj on bj.bjxh=bx.bjxh
			where xh like '#{@persion_like}%'"
		db.exec_sql(sql){|row|
			#~ p row
			hsh[row[0]]={:tch=>row[2], :name=>row[3]}
		}
		print "."
		hsh
	end

	def main db
		stds=get_stds(db)
		#~ p stds
		tch=get_tch(db)
		#~ p tch
		wd=get_wd(db)
		#~ p wd
		bj=get_bj(db)	#班级
		#~ p bj
		pj =get_pj db
		#~ p pj

		all={}
		pj.each{|key, val|
			if bj[key]
				all[bj[key][:name]] ||={}
				all[bj[key][:name]][key]=val
				#~ all<< [one[0], stds[one[0]]]
			else
				puts "Error: #{key}在班级里找不到！请检查。"
			end
		}
		#~ p all

		dj_str=['', 'A','B','C','D']

		excel = CExcel2.new
		excel.open_rw('temp.xls', 'output.xls'){|wb|
			all.sort.each{|cls, val|
				puts cls
				sht = wb.worksheets.add
				sht.name=cls
				wb.worksheets('sample').range("A1:A3").EntireRow.Copy
				sht.Paste
				sht.range('A1').value2=cls
				arr=[]
				val.sort.each{|per, others|
					sub=[]
					sub<<per
					sub<<stds[per]
					(1..3).each{|x|
						sub<<''
						(7..11).each{|y|
							[1,3,4].each{|z|
								if others[x.to_s] && others[x.to_s][y]
									sub<<dj_str[others[x.to_s][y][z].to_i]
								else
									sub<<''
								end
							}
						}
					}
					#~ others.sort.each{|nf, o1|
						#~ o1.sort.each{|pj, dj|
						#~ }
					#~ }
					arr<<sub
				}
				row_size=arr.size
				r=''
				arr<<['', '', '', '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r, '', '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r, '', '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r, '', 'A', r]
				arr<<['', '', '', '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r, '', '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r, '', '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r, '', 'B', r]

				excel.write_area sht, 'A4', arr
				#~ rg=sht.range('A4').offset(row_size)
				#~ rg=rg.offset(0, 5).FormulaR1C1 = "=COUNTIF(R4C:R57]C,\"=A\")/COUNTA(R4C:R57C)"
				
			}
			sht=wb.worksheets('sample')
			sht.delete
		}
	end
end
#~ db = Db.new({'host'=> "localhost\\sqlexpress", 'database'=> 'xx', 'username'=> 'sa', 'password'=>'sa'})
#~ db = Db.new({'host'=> "Lwang-f19de0e3c", 'database'=> 'xx', 'username'=> 'sa', 'password'=>'naominaomi'})

#~ ARGV<<'j10'
if ARGV.size==1
	$cfg = CFG.new
	$cfg.load_config
	$db = Db.new($cfg.cfg)
	all=CAll.new(ARGV[0])
	all.main($db)
else
	puts "Usage: getlist.exe 班级前3码"
	puts "e.g.: getlist.exe j10"
end
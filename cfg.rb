require "yaml"

class CFG
	attr :cfg
	attr :conn
	attr :db
	def load_config
		cfg = YAML.load_file("config.yml")
		@cfg=cfg['db']
		@db = @cfg['database']
		@conn = conn_str(cfg['db'])
	end
	def set host, db, user, pswd
		@db=db
		@conn = conn_str({'host'=>host, 'database'=>db, 'username'=>user, 'password'=>pswd})
	end
	
	protected
	def conn_str cfg_hash
		"DBI:ADO:Provider=SQLOLEDB;Connect Timeout=5;Data Source=#{cfg_hash['host']}; Initial Catalog=#{cfg_hash['database']}; Persist Security Info=False ;User ID=#{cfg_hash['username']}; Password=#{cfg_hash['password']};"
	end
end

if __FILE__==$0
	cfg = CFG.new
	#~ cfg.load_config
	cfg.set("10.191.7.56", 'IPT2Querylog', 'sa', 'sa')
	puts "Connection is #{cfg.conn}"
end

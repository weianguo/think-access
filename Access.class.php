<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK IT ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006-2014 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: weianguo <366958903@qq.com>
// +----------------------------------------------------------------------

namespace Think\Db\Driver;

use Think\Db\Driver;

/**
 * Access数据库驱动
 */
class Access extends Driver
{
    protected $selectSql = 'SELECT %LIMIT% %DISTINCT% %FIELD% FROM %TABLE%%JOIN%%WHERE%%GROUP%%HAVING%%ORDER%';
    
    /**
     * 解析pdo连接的dsn信息
     * @access public
     * @param array $config 连接信息
     * @return string
     */
    protected function parseDsn($config)
    {
        $dsn = 'odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' . $config['database'];
        return $dsn;
    }
	
	/**
     * 执行查询 返回数据集
     * @access public
     * @param string $str  sql指令
     * @param boolean $fetchSql  不执行只是获取SQL
     * @param boolean $master  是否在主服务器读操作
     * @return mixed
     */
	public function query($str, $fetchSql = false, $master = false)
    {
		if( !$fetchSql && !empty($this->config['charset']) && $this->config['charset']!='utf8'){
		   $str = iconv('UTF-8',$this->config['charset'], $str);
		}
	   
		$result = parent::query($str, $fetchSql, $master);
		if( is_array($result) && !empty($this->config['charset']) && $this->config['charset']!='utf8'){
			$result = eval('return '.iconv($this->config['charset'],'UTF-8',var_export($result,true).';'));
		}
	   return $result;
    }
	
	/**
     * 执行语句
     * @access public
     * @param string $str  sql指令
     * @param boolean $fetchSql  不执行只是获取SQL
     * @return mixed
     */
    public function execute($str, $fetchSql = false)
    {
        $this->initConnect(true);
        if (!$this->_linkID) {
            return false;
        }

        $this->queryStr = $str;
        if (!empty($this->bind)) {
            $that           = $this;
            $this->queryStr = strtr($this->queryStr, array_map(function ($val) use ($that) {return '\'' . $that->escapeString($val) . '\'';}, $this->bind));
        }
        if ($fetchSql) {
            return $this->queryStr;
        }
        //释放前次的查询结果
        if (!empty($this->PDOStatement)) {
            $this->free();
        }

        $this->executeTimes++;
        N('db_write', 1); // 兼容代码
        // 记录开始执行时间
        $this->debug(true);
		
		$sql = $this->queryStr;
	
		if( !empty($this->config['charset']) && $this->config['charset']!='utf8'){
			$sql = iconv('UTF-8',$this->config['charset'], $sql);
		}

		$this->bind = array();
        try {

            $result = $this->_linkID->exec( $sql );

            // 调试结束
            $this->debug(false);
            if (false === $result) {
                $this->error();
                return false;
            } else {
                $this->numRows = $result;
                return $this->numRows;
            }
        } catch (\PDOException $e) {
            $this->error();
            return false;
        }
    }

	
	/**
     * 获取最近一次查询的sql语句
     * @param string $model  模型名
     * @access public
     * @return string
     */
	public function getLastSql($model = '')
    {
        $str = $model ? $this->modelSql[$model] : $this->queryStr;
		if(!empty($this->config['charset']) && $this->config['charset']!='utf8'){
		   $str = iconv($this->config['charset'],'UTF-8', $str);
		}
		return $str;
    }
	
    /**
     * 取得数据表的字段信息
     * @access public
     * @return array
     */
    public function getFields($tableName)
    {
        return [];
    }

    /**
     * limit
     * @access public
     * @param $limit limit表达式
     * @return string
     */
    public function parseLimit($limit)
    {
        $limitStr = '';
        if (!empty($limit)) {
            $limit = explode(',', $limit);
            $limitStr = ' top ' . $limit[0] . ' ';
        }
        return $limitStr;
    }

     /**
     * SQL指令安全过滤
     * @access public
     * @param string $str  SQL指令
     * @return string
     */
    public function escapeString($str)
    {
        return str_ireplace("'", "''", $str);
    }
	
	/**
     * 字段和表名处理
     * @access public
     * @param string $key
     * @param bool   $strict
     * @return string
     */
    public function parseKey($key, $strict = false)
    {
        $key = trim($key);

        if ($strict && !preg_match('/^[\w\.\*]+$/', $key)) {
            E('not support data:' . $key);
        }

        if ($strict || (!is_numeric($key) && !preg_match('/[,\'\"\*\(\)`.\s]/', $key))) {
            $key = '[' . $key . ']';
        }
        return $key;
    }
}

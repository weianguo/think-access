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

    /**
     * 连接数据库方法
     * @access public
     */
    public function connect($config = '', $linkNum = 0, $autoConnection = false)
    {
        if (!isset($this->linkID[$linkNum])) {
            if (empty($config)) {
                $config = $this->config;
            }
            $file =  mb_convert_encoding($config['database'], 'GBK', 'UTF-8');
            $this->linkID[$linkNum] = new \COM("ADODB.Connection") or die("Cannot start ADO");
            $this->linkID[$linkNum]->Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=$file");
        }
        return $this->linkID[$linkNum];
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
        $this->initConnect($master);

        if (!$this->_linkID) {
            return false;
        }

        $this->queryStr = $str;
        if (!empty($this->bind)) {
            $that           = $this;
            $this->queryStr = strtr($this->queryStr, array_map(function ($val) use ($that) {
                return '\'' . $that->escapeString($val) . '\'';
            }, $this->bind));
        }
        if ($fetchSql) {
            return $this->queryStr;
        }
        //释放前次的查询结果
        if (!empty($this->PDOStatement)) {
            $this->free();
        }

        $this->queryTimes++;
        N('db_query', 1); // 兼容代码
        // 调试开始
        $this->debug(true);

        $sql = mb_convert_encoding($this->queryStr, 'GBK', 'UTF-8');
        try {
            $this->PDOStatement = $this->_linkID->execute($sql);
        } catch (\Exception $e) {
            $this->error = mb_convert_encoding($e->getMessage(), 'UTF-8', 'GBK');
            $this->error();
            return false;
        }
        $i = 0;
        $r_items = [];
        $num_fields = $this->PDOStatement->fields->count();
        while (!$this->PDOStatement->EOF) {
            for ($j = 0; $j < $num_fields; $j++) {
                $field = $this->PDOStatement->Fields($j);

                if (is_object($field->name) || is_string($field->name)) {
                    $name = mb_convert_encoding($field->name, 'UTF-8', 'GBK');
                } else {
                    $name = $field->name;
                }

                if (is_object($field->value) || is_string($field->value)) {
                    $value = mb_convert_encoding($field->value, 'UTF-8', 'GBK');
                } else {
                    $value = $field->value;
                }

                $r_items[$i][$name] = $value;
            }
            $i++;
            $this->PDOStatement->MoveNext();
        }

        return $r_items;
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
            $this->queryStr = strtr($this->queryStr, array_map(function ($val) use ($that) {
                return '\'' . $that->escapeString($val) . '\'';
            }, $this->bind));
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

        $sql = mb_convert_encoding($this->queryStr, 'GBK', 'UTF-8');
        try {
            $this->PDOStatement = $this->_linkID->execute($sql,$affected_rows);
        } catch (\Exception $e) {
            $this->error = mb_convert_encoding($e->getMessage(), 'UTF-8', 'GBK');
            $this->error();
            return false;
        }

        return $affected_rows;
    }

    /**
     * 数据库错误信息
     * 并显示当前的SQL语句
     * @access public
     * @return string
     */
    public function error()
    {
        if ('' != $this->queryStr) {
            $this->error .= "\n [ SQL语句 ] : " . $this->queryStr;
        }
        // 记录错误日志
        trace($this->error, '', 'ERR');
        if ($this->config['debug']) {
            // 开启数据库调试模式
            E($this->error);
        } else {
            return $this->error;
        }
    }
}

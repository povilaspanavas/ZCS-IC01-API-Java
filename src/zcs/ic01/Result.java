/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package zcs.ic01;

/**
 *
 * @author PPanavas
 */
public class Result {

    private byte _status;
    private int _id;
    private byte[] _serialnum;
    private byte[] _piccadata;
    
    public Result() {
    }

    void set_id(int i) {
        _id = i;
    }

    void setStatus(byte status) {
        this._status = status;
    }

    void setCardSerial(byte[] serialNum) {
        this._serialnum = serialNum;
    }

    void setResult(byte[] piccdata) {
        _piccadata = piccdata;
    }

}

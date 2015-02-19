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
public class ZCSIC01 {

    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        // copied these from API
        byte BLOCK0_EN = 0x1;
        byte BLOCK1_EN = 0x2;
        byte BLOCK2_EN = 0x4;
        byte NEEDSERIAL = 0x8;
        byte EXTERNKEY = 0x10;
        
        byte ctrword = (byte) (BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + NEEDSERIAL + EXTERNKEY);
        // I'm not sure what it is, I've put serial number of my ZCS-IC01 device
        byte[] serial = new byte[10];
        serial[0] = (byte)3;
        serial[1] = (byte)6;
        serial[2] = (byte)3;
        serial[3] = (byte)3;
        serial[4] = (byte)2;
        serial[5] = (byte)6;
        serial[6] = (byte)0;
        serial[7] = (byte)4;
        serial[8] = (byte)9;
        serial[9] = (byte)2;
        
        // this should be a default password
        byte[] password = new byte[12];
        password[0] = 0xF;
        password[1] = 0xF;
        password[2] = 0xF;
        password[3] = 0xF;
        password[4] = 0xF;
        password[5] = 0xF;
        password[6] = 0xF;
        password[7] = 0xF;
        password[8] = 0xF;
        password[9] = 0xF;
        password[10] = 0xF;
        password[11] = 0xF;
        
        
        byte area = (byte)0;
        byte[] data = new byte[50];
        data[0] = (byte)0;
        data[1] = (byte)1;
        data[2] = (byte)2;
        data[3] = (byte)3;
        data[4] = (byte)4;
        
        CLibrary library = new CLibrary();
        // Gives status 12, which means my password is incorrect
        Result myResult = getCardData(ctrword, serial, area, (byte)0, password, data);
        // Result myResult2 = writeData(ctrword, serial, area, (byte)0, password, data);
    }

    public static Result getDeviceNum(byte[] deviceNum) {
        Result rs = new Result();
        byte status;
        status = CLibrary.INSTANCE.pcdgetdevicenumber(deviceNum);
        rs.setStatus(status);
        rs.setResult(deviceNum);
        rs.set_id(1);
        return rs;
    }

    public static Result getCardData(byte ctrlword, byte[] serial, byte area,
            byte key, byte[] picckey, byte[] piccdata) {
        Result rs = new Result();
        byte status;
        byte[] serialNum = new byte[4];
        status = CLibrary.INSTANCE.piccrequest(serialNum);
        status = CLibrary.INSTANCE.piccreadex(ctrlword,
                serialNum, area, key, picckey, piccdata);
        rs.set_id(2);
        rs.setCardSerial(serialNum);
        rs.setStatus(status);
        rs.setResult(piccdata);
        return rs;
    }

    public static Result updatePass(byte ctrlword, byte[] serial, byte area, byte key,
            byte[] oldPass, byte[] newPass) {
        Result rs = new Result();
        byte status;
        byte[] serialNum = new byte[4];
        status = CLibrary.INSTANCE.piccrequest(serialNum);
        status = CLibrary.INSTANCE.piccchangesinglekey(ctrlword,
                serialNum, area, key, oldPass, newPass);
        rs.set_id(3);
        rs.setStatus(status);
        return rs;
    }

    public static Result writeData(byte ctrlword, byte[] serial, byte area, byte key, byte[] picckey, byte[] piccdata) {
        Result rs = new Result();
        byte status;
        byte[] serialNum = new byte[4];
        status = CLibrary.INSTANCE.piccrequest(serialNum);
        status = CLibrary.INSTANCE.piccwriteex(ctrlword,
                serialNum,
                area, key, picckey, piccdata);
        rs.set_id(4);
        rs.setStatus(status);
        return rs;
    }

    public static String toHex(byte[] buffer) {
        StringBuffer sb = new StringBuffer(buffer.length * 2);
        for (int i = 0; i < buffer.length; i++) {
            sb.append(Character.forDigit((buffer[i] & 240) >> 4,
                    16));
            sb.append(Character.forDigit(buffer[i] & 15, 16));
        }
        return sb.toString();
    }
}

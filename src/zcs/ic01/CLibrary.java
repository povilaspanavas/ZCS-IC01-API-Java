/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package zcs.ic01;

import com.sun.jna.Native;

/**
 *
 * @author PPanavas
 */
public class CLibrary implements ICLibrary {
    
    public static ICLibrary INSTANCE = null;

    public CLibrary()
    {
        System.setProperty("jna.library.path", "C:\\Projects\\ZCS-IC01\\build\\classes\\zcs\\ic01");
        INSTANCE = (ICLibrary) Native.loadLibrary("OUR_MIFARE", ICLibrary.class);
    }
    public byte pcdgetdevicenumber(byte[] deviceNum) {
        return INSTANCE.pcdgetdevicenumber(deviceNum);
    }

    /**
     * n @return d card l serial number
     */
    public byte piccrequest(byte[] serial) {
        return INSTANCE.piccrequest(serial);
    }

    /**
     * n @return d read k block 0 0 0 0~2
     */
    public byte piccreadex(byte ctrlword, byte[] serial, byte area, byte key, byte[] picckey, byte[] piccdata) {
        return INSTANCE.piccreadex(ctrlword, serial, area, key, picckey, piccdata);
    }

    /**
     * n @return e change password
     */
    public byte piccchangesinglekey(byte ctrlword, byte[] serial,
            byte area, byte key, byte[] oldPass, byte[] newPass) {
        return INSTANCE.piccchangesinglekey(ctrlword, serial, area, key, oldPass, newPass);
    }

    /**
     * n @return write
     */
    public byte piccwriteex(byte ctrlword, byte[] serial, byte area, byte key, byte[] picckey, byte[] piccdata) {
        return INSTANCE.piccwriteex(ctrlword, serial, area, key, picckey, piccdata);
    }

    /**
     * m @param buffer n @return h switch byte[] to string
     */
    public String toHex(byte[] buffer) {
        return INSTANCE.toHex(buffer);
    }
}

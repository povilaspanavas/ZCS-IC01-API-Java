/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package zcs.ic01;

import com.sun.jna.Native;
import com.sun.jna.win32.StdCallLibrary;

/**
 *
 * @author PPanavas
 */
public interface ICLibrary extends StdCallLibrary {


                // System.setProperty("C:\\Projects\\ZCS-IC01\\build\\classes\\zcs\\ic01", "OUR_MIFARE");
    /**
     * n @return e device l serial number
     */
    public byte pcdgetdevicenumber(byte[] deviceNum);

    /**
     * n @return d card l serial number
     */
    public byte piccrequest(byte[] serial);
    

    /**
     * n @return d read k block 0 0 0 0~2
     */
    public byte piccreadex(byte ctrlword, byte[] serial, byte area, byte key, byte[] picckey, byte[] piccdata);

    /**
     * n @return e change password
     */
    public byte piccchangesinglekey(byte ctrlword, byte[] serial,
            byte area, byte key, byte[] oldPass, byte[] newPass);

    /**
     * n @return write
     */
    public byte piccwriteex(byte ctrlword, byte[] serial, byte area, byte key, byte[] picckey, byte[] piccdata);

    /**
     * m @param buffer n @return h switch byte[] to string
     */
    public String toHex(byte[] buffer);
}

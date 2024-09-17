
//import com.mitac.common.support.uil.TypeConversions;
import io.jsonwebtoken.Jwts;
import io.jsonwebtoken.SignatureAlgorithm;
import org.junit.Test;

import java.security.*;
import java.security.spec.InvalidKeySpecException;
import java.security.spec.KeySpec;
import java.security.spec.PKCS8EncodedKeySpec;
import java.util.Base64;
import java.util.UUID;
import java.util.concurrent.TimeUnit;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class TokenGenerator {

    /**
     * 生產公私鑰
     *
     * @throws NoSuchAlgorithmException
     * @throws InvalidKeySpecException
     */
    public static void genKey() throws Exception {

        // 生產公私鑰
        KeyPairGenerator keyPairGenerator = KeyPairGenerator.getInstance("RSA");
        keyPairGenerator.initialize(2048, new SecureRandom(UUID.randomUUID().toString().getBytes()));

        KeyPair keyPair = keyPairGenerator.generateKeyPair();

        String privateKey = Base64.getEncoder().encodeToString(keyPair.getPrivate().getEncoded());
        String publicKey = Base64.getEncoder().encodeToString(keyPair.getPublic().getEncoded());

        System.out.println("PRIVATE KEY: " + privateKey);
        System.out.println("PUBLIC KEY: " + publicKey);
        
        //privateKey = "MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCH1iotdMtusCcudzP4lfdcxBiJiU32mXWCoLvxFgEmLO3uW4uTz1FKNZ_4-VvHO-4hE-UtwPF0IVBsad1fzoOTCCJ8yG7nVR8ZStlc68uWg9mUC5yyUkFgIRqmgJktHjuw1jCE-MseoM3faz5FxBXdbAJlUDlYmK0CZJXA2RooRQnE78YCLwlnyUo2-ClOMoF_CUtz5mlhyJNnbrVlZb3CueaL4FJqdmw0QwhzDz8736sgAjDEDDUjZv9ogF1MfbNmNhykClXmG_HSpqmegzeaHX941JPJ-J_TqhOyPgF47-smKhKJu3muQydBOWr1Ea4q1FnWLnozSknDxJgYUDCNAgMBAAECggEAcWRCSTfaKkg6LPutErJ9j57SrN1Fi7mG8siimxo3U1rmM7ePyI-j2ELzi679Ak_w9QPaqFsMNFkq_ZVSCwwlobOEto3KpqnUEBT_ZiYgCUF_e6pF4EXx9QEtchifxZ4bTf8--YGGbcbmlL69eRe6-N-VEGXruR2aLwkwSY_x2fKwYQxOCTz2BNrxc6vS-UBONyWQ1UZvklxgCTYPWmLsYdzhKW6dU0w-IENjQ6WCWfzwGYo4Wjgi7CchB1u0SKMgxjA2GaUamoJJ5FAvM3JaE_aaqAYWFwSgDwbfTfeF0Esa8k6yZOgIDoIYhFAhCxcNPbA2zDPCvei68hSBNTWHMQKBgQC7zOUOXY4MwRk1o3aE8Egt06O2WJ-qq08-auwGwh52r_LU2O4m3kWl5y9VbhccImUWYVmst6oT8XP7jUqj_dtvHbotsNNUBa9lnwFv5yKAtkJKVZGK3sK-aJP_BdS1L7Yc5mK09ZOJBOvKB0SN-XzCzGAEDj2r98V6USTaPY3V-wKBgQC5KmOOHs2TWzub5FKmC-Ex9jY-77SOkUhMqxkYA_s75THIULxrSmjiGCL5m5rDvGZWbCPzAQT8W4x-VROQ5PURDApfZX1AzpAz5huKmLrZ3MJWoT8JwYOGeKlW7gDMhVegVA-X4sImRoCkFGy-KJjUb74eDpGYJlB0DjT2F781FwKBgQCGIN3Lt8fXliaCJ8BjTBXRHSIE_yDyTS3ov4bZgNUvIZVGrbTR79hAmHA3DMzWnD4kkNzyPa7sVXvnws73dzy9DLdHIM6eaP0PkFP_b042LXYFDz5Gt7jRM3HYJ3r2-R-RXn5LDkYUC364KR6uY-zWWA-PdfhYFTtmlAPFF0dw9wKBgQCZf7tiIMT5CGOk-yVMw4JfAaW8jMhYe7W84QX_c6V85KZdUhiwtNG3xJyR4d3tr6wCrskqdMjmTxprzClZL4S9KgcbSC4KYHMIoxRn0-7qFmkAmdGBS-u1uSdgihMdeNjIb4cxuWiLhFy2KLxw84Smby_jCN7Hi9OcMf7Tl6IJ5wKBgCLGGkgTA6nH55bzwDKQapEi6DpsYAZ6T5QbzVLIve47AJmByDnec3ZypwEk_UlI1JtSKg4MKj3C42pKqFQmb1E6pQ_nZFA8qrwKLVMkXK3YLqURebKoga8L0NSyPgFa9Dfb3vK6z-ooSA5YfgQyz44_-XXH7uUBROtpHlyt5dhJ";
        encryptKey(privateKey);

    }

    /**
     * 私鑰加密
     *
     * @param privateKey
     * @throws NoSuchAlgorithmException
     * @throws InvalidKeySpecException
     */
    public static void encryptKey(String privateKey) throws Exception {

        // 進行私鑰加密
        KeyFactory keyFactory = KeyFactory.getInstance("RSA");
        byte[] priKeyBytes = Base64.getDecoder().decode(privateKey);
        KeySpec keySpec = new PKCS8EncodedKeySpec(priKeyBytes);
        PrivateKey priKey = keyFactory.generatePrivate(keySpec);
        
        Date now = new Date();
        System.out.println("michelle now=" + now.getTime());
        TimeUnit.MILLISECONDS.toSeconds(now.getTime());



        
    	DateFormat dateFormat2 = new SimpleDateFormat("yyyyMMddHHmmss");
    	Date myDate2 = dateFormat2.parse("20230831114909");        
        
        String tokenContent = Jwts.builder()
                .claim("interchangeId", "20230831114909") // 根據request "interchangeId"
                .setIssuer("NB") // 根據request "sender"
                //.setIssuedAt(TypeConversions.parseFromLocalDatetime(20230831114909L)) // 根據request "timeIsSentAt"
                //.setIssuedAt(now) // 根據request "timeIsSentAt"
                .setIssuedAt(myDate2) // 根據request "timeIsSentAt"
                .signWith(priKey, SignatureAlgorithm.RS256) //演算法
                .compact();

        System.out.println("ENCRYPTED PRIVATE KEY: " + tokenContent);
    }
    
    public static void main(String[] args) throws Exception {
    	genKey();

    }
}

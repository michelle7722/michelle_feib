
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.KeyFactory;
import java.security.KeyPair;
import java.security.KeyPairGenerator;
import java.security.PrivateKey;
import java.security.Signature;
import java.security.interfaces.RSAPrivateKey;
import java.security.interfaces.RSAPublicKey;
import java.security.spec.PKCS8EncodedKeySpec;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;


public class RSAUtils {
	
	public static final String SIGNATURE_ALGORITHM = "SHA256withRSA";
	
	public static Map<String,byte[]> keyMap = new HashMap<String, byte[]>();
	
	public static void main(String[] args) {
		RSAUtils.genKey("c:\\key\\public.key", "c:\\key\\private.key");
	}
	
	/**
	 * 產生公私鑰
	 * @param publicKeyPath
	 * @param privateKeyPath
	 */
	public static void genKey(String publicKeyPath, String privateKeyPath) {
		FileOutputStream publicFos = null;
		FileOutputStream privateFos = null;
		try {
			KeyPairGenerator keyPairGenerator = KeyPairGenerator.getInstance("RSA"); 
			keyPairGenerator.initialize(2048); 
			KeyPair keyPair = keyPairGenerator.generateKeyPair();
			RSAPublicKey rsaPublicKey = (RSAPublicKey)keyPair.getPublic();
			RSAPrivateKey rsaPrivateKey = (RSAPrivateKey)keyPair.getPrivate();
			publicFos = new FileOutputStream(publicKeyPath);
			byte[] bArray = rsaPublicKey.getEncoded();
					
			publicFos.write(Base64.getUrlEncoder().withoutPadding().encodeToString(rsaPublicKey.getEncoded()).getBytes());
			privateFos = new FileOutputStream(privateKeyPath);
			//privateFos.write(Base64.getEncoder().withoutPadding().encodeToString(rsaPublicKey.getEncoded()).getBytes());
			privateFos.write(Base64.getUrlEncoder().withoutPadding().encodeToString(rsaPrivateKey.getEncoded()).getBytes());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (publicFos != null) {
				try {
					publicFos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (privateFos != null) {
				try {
					privateFos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 以私鑰加上簽章
	 * @param privateKey
	 * @param text
	 * @return
	 * @throws Exception
	 */
	public static String sign(PrivateKey privateKey, String text) throws Exception {
		Signature sign = Signature.getInstance(SIGNATURE_ALGORITHM);
        sign.initSign(privateKey);
        sign.update(text.getBytes());
        byte[] signed = sign.sign();
        return Base64.getUrlEncoder().withoutPadding().encodeToString(signed);
	}
	
	/**
	 * 取得Key資料
	 * @param keyPath
	 * @return
	 * @throws Exception
	 */
	public static byte[] getKey(String system, String keyPath) throws Exception{
		byte[] key = keyMap.get(system);
		if (key == null) {
			FileInputStream keyfis = new FileInputStream(keyPath);
			key = new byte[keyfis.available()];
			keyfis.read(key);
			keyfis.close();
			key = Base64.getUrlDecoder().decode(key);
			keyMap.put(system, key);
		}
		return key;
	}
	
	/**
	 * 以私鑰加上簽章
	 * @param privateKeyPath
	 * @param text
	 * @return
	 * @throws Exception
	 */
	public static String signByPath(String system, String privateKeyPath, String text) throws Exception{
		byte[] privateKeyBytes = getKey(system,privateKeyPath);
		PKCS8EncodedKeySpec keySpec = new PKCS8EncodedKeySpec(privateKeyBytes);
		KeyFactory keyFactory = KeyFactory.getInstance("RSA");
		RSAPrivateKey priKey = (RSAPrivateKey)keyFactory.generatePrivate(keySpec);
		return sign(priKey, text);
	}
}

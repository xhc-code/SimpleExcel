package cn.dream.excep;

/**
 * 未声明 {@code Excel} 注解
 * @author Dream
 *
 */
public class NotDeclareExcelAnnoException extends ExcelRuntimeException {

	public NotDeclareExcelAnnoException(String msg) {
		super(msg);
	}

}

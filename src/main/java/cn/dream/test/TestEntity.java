package cn.dream.test;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class TestEntity {

	public static String staticProp;

	public static final String staticFinalProp = "";

	public final String finalProp = "";
	
	private String commonProp;
	
}

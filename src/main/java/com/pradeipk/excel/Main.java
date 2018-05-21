package com.pradeipk.excel;

/**
 * @pradeipk Created 20/05/2018
 *
 */
public class Main {
	public static void main(String[] args) {

		GenerateMemberCode generateMemberCode = new GenerateMemberCode();
		try {
			generateMemberCode.generateMemberCode();
		}
		catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}
}

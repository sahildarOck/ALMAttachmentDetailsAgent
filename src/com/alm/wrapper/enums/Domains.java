package com.alm.wrapper.enums;

public enum Domains {

	DOMAIN_NAME1("Domain Name1"),
	DOMAIN_NAME2("Domain Name2");
	
	private String domain;
	
	Domains(String domain) {
		this.domain = domain;
	}
	
	public String getDomain() {
		return domain;
	}
}

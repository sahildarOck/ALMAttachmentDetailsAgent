package com.alm.wrapper.enums;

public enum ExecutionStatus {
PASSED("Passed"),
FAILED("Failed"),
BLOCKED("Blocked");
	
	private String status;
	
	ExecutionStatus(String status) {
		this.status = status;
	}

	public String getStatus() {
		return status;
	}
}
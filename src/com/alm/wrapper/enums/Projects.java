package com.alm.wrapper.enums;

public enum Projects {

	PROJECT_NAME1("Project Name1"),
	PROJECT_NAME2("Project Name2");
	
	private String project;
	
	Projects(String project) {
		this.project = project;
	}
	
	public String getProject() {
		return project;
	}
}

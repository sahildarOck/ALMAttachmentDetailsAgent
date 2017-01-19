# ALMAttachmentDetailsAgent
Project to fetch attachment details and perform some validations on them

An automation tool developed to address the requirement of performing audit of ensuring that the attachments linked with the test cases in ALM's Test Lab conforms the conventions.
It takes either test set id or test folder name as input and outputs all the attchments in them in an excel with other required details like, the convention compliance, attachment size, test id, tester, test case name etc.

It uses Java Swings API for the front-end and the OTA COM API of HP for back-end.

The entire coding of the tool has been done using Java and JACOB (Java Com Bridge) to make JNI calls to the COM API of HP-ALM.

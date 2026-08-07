package com.bada.hello_spring.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class HelloController {
	
	// 정적페이지 반환
	@GetMapping("/hello")
	public String hello (Model model) {
		model.addAttribute("data", "hello!!!");
		return "hello";
	}
	
	// mvc구조로 반환
	@GetMapping("/hello-mvc")
	public String helloMvc (@RequestParam("name") String name, Model model) {
		model.addAttribute("name", name);
		return "hello-template";
	}
	
	// api형태로 반환
	@GetMapping("/hello-api")
	@ResponseBody
	public String helloApi (@RequestParam(value = "name", required = false) String name) {
		// model에 데이터를 담을 필요가 없다.
		// @ResponseBody 애너테이션을 선언하게 되면, 뷰리졸버가 아닌 httpmessageconverter쪽으로 던진다.
		// 응답 http의 body에 리턴하는 데이터(문자열)을 담아서 브라우저에 그대로 표시한다.
		// model.addAttribute("name", name);
		return "바다 & " + name;
	}
	
	// responsebody 객체를 반환시켜보기
	@GetMapping("/hello-api-test")
	@ResponseBody
	public Bada helloApiTest (@RequestParam String address) {
		Bada bada = new Bada();
		bada.setAddress((address == null || address.isEmpty()) ? "대구시": address);
		// @ResponseBody를 선언하고 객체를 반환하게 되면, 뷰리졸버가 아닌 httpmessageconverter가 동작하여, JSON 형태로 변환하여 브라우저에 반환
		return bada;
	}
	
	static class Bada {
		private String name = "이바다";
		private String address;
		
		public String getAddress() {
			return address;
		}
		
		public void setAddress(String address) {
			this.address = address;
		}
		
		public String getName() {
			return name;
		}
		
		public void setName(String name) {
			this.name = name;
		}
	}
	
	
}

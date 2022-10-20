package com.org.combineexel;

import com.org.combineexel.util.MergeMultipleXlsFilesInDifferentSheet;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class CombineexelApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(CombineexelApplication.class, args);
		MergeMultipleXlsFilesInDifferentSheet.mergeExcelFiles();
	}

}

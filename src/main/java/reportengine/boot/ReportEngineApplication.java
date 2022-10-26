package reportengine.boot;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import reportengine.core.ReportWritter;
import reportengine.data.ObjectDataDTO;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class ReportEngineApplication {

	static ObjectDataDTO data = new ObjectDataDTO("Alisson", 25, 1.77f);
	static List<ObjectDataDTO> dataValues = new ArrayList<>();

	public static void main(String[] args) throws IOException, IllegalAccessException {
		SpringApplication.run(ReportEngineApplication.class, args);
		ReportWritter reportWritter = new ReportWritter<>();
		for(int i = 0; i < 10; i++) {
			dataValues.add(data);
		}
		reportWritter.writeInXlsx(data, dataValues);
	}

}

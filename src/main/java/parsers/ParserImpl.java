package parsers;

import model.CadNumber;
import org.openqa.selenium.WebDriver;

import java.util.List;
import java.util.Set;

public interface ParserImpl {

    List<CadNumber> getListCad(Set<String> cadNum, WebDriver driver);
}

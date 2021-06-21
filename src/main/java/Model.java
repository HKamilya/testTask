import lombok.Data;

import java.util.Date;

@Data
public class Model {
    private double low;
    private double high;
    private double upperBand;
    private double lowerBand;
    private Date dateTime;
    private Date dateTimeForHigh;
    private Date dateTimeForLow;
}

package members;

import java.io.IOException;

import org.springframework.stereotype.Component;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

@Component
public class HotelEntity {

	private long id;
	private String name;
	private String city;
	private long checkin;
	private long checkout;
	private double price;
	private double fullPrice;
	private int occupancy;  // for how many persons room fits
	private DealType dealType;
	private Currency currency;
	private RoomType roomType;
	private SOURCE source;
	private boolean refundable;
	private boolean plan;  // Breakfast
	private boolean exclusive;
	private float rating;
	
	public long getId() {
		return id;
	}
	
	public void setId(long id) {
		this.id = id;
	}
	
	public String getName() {
		return name;
	}
	
	public void setName(String name) {
		this.name = name;
	}
	
	public String getCity() {
		return city;
	}
	
	public void setCity(String city) {
		this.city = city;
	}
	
	public long getCheckin() {
		return checkin;
	}
	
	public void setCheckin(long checkin) {
		this.checkin = checkin;
	}
	
	public long getCheckout() {
		return checkout;
	}
	
	public void setCheckout(long checkout) {
		this.checkout = checkout;
	}
	
	public DealType getDealType() {
		return dealType;
	}
	
	public void setDealType(DealType dealType) {
		this.dealType = dealType;
	}
	
	public double getPrice() {
		return price;
	}
	
	public void setPrice(double price) {
		this.price = price;
	}
	
	public double getFullPrice() {
		return fullPrice;
	}
	
	public void setFullPrice(double fullPrice) {
		this.fullPrice = fullPrice;
	}
	
	public Currency getCurrency() {
		return currency;
	}
	
	public void setCurrency(Currency currency) {
		this.currency = currency;
	}
	
	public RoomType getRoomType() {
		return roomType;
	}
	
	public void setRoomType(RoomType roomType) {
		this.roomType = roomType;
	}
	
	public int getOccupancy() {
		return occupancy;
	}
	
	public void setOccupancy(int occupancy) {
		this.occupancy = occupancy;
	}
	
	public boolean isRefundable() {
		return refundable;
	}
	
	public void setRefundable(boolean refundable) {
		this.refundable = refundable;
	}
	
	public boolean isPlan() {
		return plan;
	}
	
	public void setPlan(boolean plan) {
		this.plan = plan;
	}
	
	public SOURCE getSource() {
		return source;
	}
	
	public void setSource(SOURCE source) {
		this.source = source;
	}
	
	public boolean isExclusive() {
		return exclusive;
	}
	
	public void setExclusive(boolean exclusive) {
		this.exclusive = exclusive;
	}
	
	public float getRating() {
		return rating;
	}
	
	public void setRating(float rating) {
		this.rating = rating;
	}
	
	/* (non-Javadoc)
	 * @see java.lang.Object#toString()
	 */
	@Override
	public String toString() {
		return "HotelEntity [id=" + id + ", name=" + name + ", city=" + city + ", checkin=" + checkin + ", checkout="
				+ checkout + ", price=" + price + ", fullPrice=" + fullPrice + ", occupancy=" + occupancy
				+ ", dealType=" + dealType + ", currency=" + currency + ", roomType=" + roomType + ", source=" + source
				+ ", refundable=" + refundable + ", plan=" + plan + ", exclusive=" + exclusive + ", rating=" + rating
				+ "]";
	}

	public static void main(String[] args) {
		
		ObjectMapper mapper = new ObjectMapper();
		HotelEntity hotelEntitiy = new HotelEntity();
		
		try {
			// Convert object to JSON string and save into a file directly
			//mapper.writeValue(new File("c:\\staff.json"), hotelEntitiy);

			// Convert object to JSON string
			String jsonInString = mapper.writeValueAsString(hotelEntitiy);
			System.out.println(jsonInString);

			// Convert object to JSON string and pretty print
			jsonInString = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(hotelEntitiy);
			System.out.println(jsonInString);

		} catch (JsonGenerationException e) {
			e.printStackTrace();
		} catch (JsonMappingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
}

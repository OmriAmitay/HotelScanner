package parser;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Random;

public class MockUtils {

	public static List<HotelEntity> getMocks(int size) {
		
		List<HotelEntity> mocks = new ArrayList<HotelEntity>(size);
		mocks.add(getSingleMock());
		return mocks;
		
	}
	
	private static HotelEntity getSingleMock() {
		
		Random rnd = new Random();
		HotelEntity entity = new HotelEntity();
		
		entity.setName("Ramada");
		entity.setCity("Natania");
		getCheckInTime(entity);
		entity.setDealType(DealType.HB);
		entity.setFullPrice(1000d);
		entity.setPrice(Math.floor(getPrice(rnd) * entity.getFullPrice()));
		entity.setCurrency(Currency.NIS);
		entity.setRoomType(RoomType.REGULAR);
		entity.setOccupancy(2);
		entity.setRefundable(false);
		entity.setPlan(true);
		entity.setSource(SOURCE.ISSTA);
		entity.setExclusive(false);
		entity.setRating(getRating(rnd));
		
		return entity;
		
	}

	private static void getCheckInTime(HotelEntity entity) {
		Calendar calendar = Calendar.getInstance();
		entity.setCheckin(calendar.getTimeInMillis());
		calendar.add(Calendar.DAY_OF_YEAR, 1);
		entity.setCheckout(calendar.getTimeInMillis());
	}
	
	private static float getRating(Random rnd) {
		float minX = 50.0f;
		float maxX = 100.0f;
		float randomFloat = rnd.nextFloat() * (maxX - minX) + minX;
		return randomFloat;
	}
	
	private static double getPrice(Random rnd) {
		double minX = 50.0f;
		double maxX = 100.0f;
		double randomDouble = rnd.nextDouble() * (maxX - minX) + minX;
		return randomDouble / maxX;
	}
	
	public static void main(String[] args) {
		List<HotelEntity> mocks = MockUtils.getMocks(1);
		System.out.println(mocks.get(0));
	}
	
}

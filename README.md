# Excel DownLoad

## Documentation
MyBatis ResultHandler 와 Fetch Size 조정을 통한 대용량 Excel File Dwonload

 ## Spec
  - Java 8
  - MyBatis + Mysql
  - Spring boot + Gradle
  - Apache POI

## 1. MyBatis Result Handler  

#### org.apache.ibatis.session.ResultHandler<T>  

##### Public Methods
    public abstract void handleResult (ResultContext<? extends T> resultContext)
  


####  1.1 ResultHandler의 handleResult 구현

###### Service.java
```java
@Override
public void setBody() {
    // ResultHandler를 생성, write 메소드를 사용할 수 있도록 넘겨준다.
    ResultHandler<Map<String,Object>> resultHandler = this::write;
    downMapper.get(resultHandler);
}

@Override
protected void write(ResultContext<?> resultContext) {
    Map<String,Object> map = (Map<String,Object>) resultContext.getResultObject();
    rowNum++;
    SXSSFRow row = createRow(1); // 신규 Row 생성.
    createCell(row, 1, String.valueOf(rowNum)); // No 작성.
    AtomicInteger cellNo = new AtomicInteger(1);

    map.forEach((k,v)-> createCell(row, cellNo.getAndIncrement(), k  +"_"+ v)); // Cell+1 씩하면서 Map에 들은 Data작성
    // ROW가 6만건 이상 일시, 신규시트를 생성하여, 신규시트에 작성
    if (getRows() > MAX_EXCEL_ROW_SIZE) {
        createSheet(); // 신규 시트 생성
    }
}
  ```
##### Row단위로 Data를 컨트롤 할 수 있게 해주며, 수많은 List를 메모리에 담지않고, Fetch Size 만큼의 Row Data를 통제할 수 있도록 지정해준다.
##### 추가로, 6만건 이상 일 경우 신규시트를 작성하여, 6만건 + @의 Data를 기록할 수 있도록 처리하자. 

####  1.2 Mapper interface 작성.
###### Mapper.java
```java
@Mapper
public interface DownMapper {
    void get(ResultHandler<Map<String, Object>> resultHandler); // void로 선언
}
```
##### Mapper에서는 ResultHandler를 넘기고, 리턴타입은 꼭 void로 선언하자.
  
>https://mybatis.org/mybatis-3/apidocs/reference/org/apache/ibatis/session/ResultHandler.html
  * * *
  
## 2. Fetch Size
####  2.1 xml SQL 작성.
###### dwon.xml
```sql
<select id="get" resultType="Map" fetchSize="1000">
    SELECT NAME, ID, DEPT
    FROM USER
</select>
```
  ####  Fetch Size 적용하여, 서버성능에 맞게 테스트하면서, Size를 조절하자. (높을수록 빠르다.)

> VO 등으로 resultMap을 구현할 경우 Data Mapping시 속도가 느려진다.
  
  ## 3. Mysql useCursorFetch=true 옵션 적용
####  fetchSize를 적용하긴 위해서 Mysql 해당옵션을 활성화 시켜줘야한다.  
###### application.properties.   
```yaml
spring.datasource.hikari.jdbc-url=jdbc:mysql://localhost:3376/NAME?useCursorFetch=true
```
  ##### 정상적용시 RowDataStatic에서 RowDataCursor 로 클래스가 변경된것을 볼 수 있다.

    
## 3. 결론
  약, 15개의 컬럼과 30만건의 Row 테스트시 무리없이 생성되었으나, Excel 용량이 너무 많아져 Excel이 오픈하는게 느리다.. 10만건정도가 최선인듯

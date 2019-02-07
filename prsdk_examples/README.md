# prsdk_examples
Some Pentaho Reporting Engine SDK Examples

Using the Pentaho Reporting Engine SDK to enable and create extensions and functions.

## Dependencies

To be able to use the Pentaho Report Engine SDK:

```
		<dependency>
			<groupId>org.pentaho.reporting.engine</groupId>
			<artifactId>classic-core</artifactId>
			<version>8.2.0.0-SNAPSHOT</version>
			<scope>compile</scope>
		</dependency>
```
To be able to use the Pentaho Report Engine SDK Samples:

```
		<dependency>
			<groupId>org.pentaho.reporting.engine</groupId>
			<artifactId>classic-samples</artifactId>
			<version>8.2.0.0-SNAPSHOT</version>
		</dependency>
```

## First Example: Cronjob

For cronjob we're using Spring:

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter</artifactId>
		</dependency>
```
To enable the cronjob we have to annotate the class which contains the main method

```
  @SpringBootApplication
  @EnableScheduling
```
and we have to annotate the method which sould be scheduled

```
@Scheduled(fixedRate = 10000)
```
the arguement "fixedRate" defines the intervals on which the method will called.

At the method "generate" in the "SampleReportGenerator" class, we define the file name and the output type.
At the method "getReportDefinition", we define the masterreport file (prpt).

I got errors caused by springs logging factory. To fix this:

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter</artifactId>
			<exclusions>
				<exclusion>
					<groupId>org.springframework.boot</groupId>
					<artifactId>spring-boot-starter-logging</artifactId>
				</exclusion>
			</exclusions>
		</dependency>
```

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter-log4j</artifactId>
			<version>1.3.8.RELEASE</version>
		</dependency>
```

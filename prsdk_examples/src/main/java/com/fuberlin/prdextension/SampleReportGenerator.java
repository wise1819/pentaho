package com.fuberlin.prdextension;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Map;

import org.pentaho.reporting.engine.classic.core.DataFactory;
import org.pentaho.reporting.engine.classic.core.MasterReport;
import org.pentaho.reporting.engine.classic.core.ReportProcessingException;
import org.pentaho.reporting.engine.classic.samples.AbstractReportGenerator;
import org.pentaho.reporting.libraries.resourceloader.Resource;
import org.pentaho.reporting.libraries.resourceloader.ResourceException;
import org.pentaho.reporting.libraries.resourceloader.ResourceManager;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

@Component
public class SampleReportGenerator extends AbstractReportGenerator{

	@Override
	public DataFactory getDataFactory() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public MasterReport getReportDefinition() {
	    try {
	        // Using the classloader, get the URL to the reportDefinition file
	        final ClassLoader classloader = this.getClass().getClassLoader();
	        final URL reportDefinitionURL = classloader.getResource( "SampleReportCrossTab.prpt" );

	        // Parse the report file
	        final ResourceManager resourceManager = new ResourceManager();
	        final Resource directly = resourceManager.createDirectly( reportDefinitionURL, MasterReport.class );
	        return (MasterReport) directly.getResource();
	      } catch ( ResourceException e ) {
	        e.printStackTrace();
	      }
	      return null;
	}

	@Override
	public Map<String, Object> getReportParameters() {
		// TODO Auto-generated method stub
		return null;
	}
	
	@Scheduled(fixedRate = 10000)
	public void generate() throws IllegalArgumentException, IOException, ReportProcessingException {
	    final File outputFilename = new File( "Test.pdf" );

	    // Generate the report
	    new SampleReportGenerator().generateReport( AbstractReportGenerator.OutputType.PDF, outputFilename );

	    // Output the location of the file
	    System.err.println( "Generated the report [" + outputFilename.getAbsolutePath() + "]" );
	}

	
}

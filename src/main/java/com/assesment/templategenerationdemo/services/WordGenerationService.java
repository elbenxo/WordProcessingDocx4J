package com.assesment.templategenerationdemo.services;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.docx4j.XmlUtils;
import org.docx4j.finders.TableFinder;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import jakarta.xml.bind.JAXBElement;
import lombok.extern.log4j.Log4j2;

/*
 * This service helps create a word from a template. 
 */
@Service
@Log4j2
public class WordGenerationService {

    @Value("${template.templateUrlString}")
    private String templateUrlString;
    @Value("${template.pdfPathUrl}")
    private String pdfPathUrl;

    public void executeOperation() {

        WordprocessingMLPackage wordMLPackage = readTemplate(templateUrlString);
        

        if (wordMLPackage == null) {
            log.info("Template not found");
            return;
        }

        var variables = generateVariables();

		try {

			VariablePrepare.prepare(wordMLPackage);
			
			MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
	        log.info("XML: {}", mainDocumentPart.getXML());

			/*
			 * Obtenemos los nombres de tablas que tienen Parametros. Empiezan por TableX.
			 * siendo X un identificador de tabla unico que se define en: Table properties->
			 * Pestaña "Alt text" y campo "Title".
			 */
	        log.info("Busqueda de tablas con parametros");
			List<String> tablas = getTableNamesToParametrized(variables);

			log.info("Tablas con parametros encontradas: {}", tablas);
			
			for (String t : tablas) {
				var tabla = findTableById(mainDocumentPart, t);
				if (tabla != null) {
					log.info("Tabla '{}' encontrada en el documento", t);
					generateLines(tabla, getNumberofLines(t, variables));
				}else {
					log.info("Tabla '{}' no encontrada en el documento", t);
				}
			}

			// Prueba de que se puede borrar una Tabla.
			var tabla2 = findTableById(mainDocumentPart, "Tabla2");
			removeTableFromDocument(tabla2);

			// Sustitucion de todas las variables.
			mainDocumentPart.variableReplace(generateVariables());
		} catch (Exception e) {
			e.printStackTrace();
		}
        
        //Grabamos a disco el nuevo documento Word.
        saveToDisk(wordMLPackage, pdfPathUrl);

    }
    
    /**
     * Obtiene el numero de lineas de una tabla por el sufijo de sus propiedades.
     * @param tableName Nombre de la tabla.
     * @param variables Propiedades.
     * @return Numero maximo de lineas segun las variables definidas.
     */
    private int getNumberofLines(String tableName, Map<String, String> variables) {
    	
    	return variables.keySet().stream().filter(i-> i.startsWith(tableName))
    					 .map(i -> i.substring(i.lastIndexOf("_") +1 , i.length()))
    					 .distinct()
    					 .mapToInt(v -> Integer.valueOf(v) + 1)
    					 .max()
    					 .orElse(0);

    }
    
    /**
     * Obtiene el nombre de las tablas que tienen propiedades para modificar.
     * La propiedad y la tabla tienen que empezar por "TableX." 
     * 
     * @param variables variables de parametrizacion.
     * @return Listado de nombres de tablas.
     */
    private List<String> getTableNamesToParametrized(Map<String, String> variables){
    	
    	List<String> tablas = Collections.emptyList();
    	
    	if (variables != null && variables.keySet() != null) {
    		tablas = variables.keySet().stream()
    						.filter(i -> i.startsWith("Table")).map(i-> i.substring(0, i.indexOf(".")))
    						.distinct()
    						.collect(Collectors.toList());
    	}
    	
    	return tablas;
    	
    }
    
    
    /**
     * Generacion de Variables a modificar de ejemplo.
     * @return un Hashmap con los valores de las variables de la plantilla a completar.
     */
    private Map<String, String> generateVariables(){
    	var variables = new HashMap<String, String>();
    	
    	variables.put("Customer Name", "Federico Jimenez losantos");
    	variables.put("Customer Address", "Calle Gran Via 1 1ºA 28001 Madrid");
    	variables.put("Customer Phone", "+34 91 6665544");
    	variables.put("Customer Email", "info@soyou.es");
    	variables.put("Credit Card Number", "4003-9010-5106-7773");
    	variables.put("Credit Card Type", "Credito");
    	variables.put("Credit Card Limit", "150 €");
    	variables.put("Start Date", "01/01/2025");
    	variables.put("End Date", "01/01/2027");
    	variables.put("Total Purchases", "1000€");
    	variables.put("Total Payments", "2000€");
    	variables.put("Total Interest", "5%");
    	variables.put("Total Fees", "250€");
    	variables.put("Previous Balance", "0€");
    	variables.put("New Balance", "250€");
    	variables.put("Minimum Payment Due", "100€");
    	variables.put("Payment Due Date", "01/02/2025");
    	variables.put("Transaction Date", "04/01/2025");
    	variables.put("Transaction Description", "Transaccion");
    	variables.put("Transaction Amount", "222.11€");
    	variables.put("Transaction Balance", "322.11€");
    	variables.put("Customer Service Phone", "+34 659887766");
    	variables.put("Customer Service Email", "cliente@cliente.com");
    	variables.put("Website URL", "http://soyou.es");

    	variables.put("Table1.Column1Value_0", "Pepito");
    	variables.put("Table1.Column2Value_0", "Perez");
    	variables.put("Table1.Column3Value_0", "Gonzalez");
    	variables.put("Table1.Column4Value_0", "50.000");

    	variables.put("Table1.Column1Value_1", "Juanito");
    	variables.put("Table1.Column2Value_1", "Muñoz");
    	variables.put("Table1.Column3Value_1", "Rodriguez");
    	variables.put("Table1.Column4Value_1", "75.000");



    	variables.put("Table1.Column1Value_2", "Fulanito");
    	variables.put("Table1.Column2Value_2", "Menganez");
    	variables.put("Table1.Column3Value_2", "Jimenez");
    	variables.put("Table1.Column4Value_2", "45.000");

    	
    	return variables;
    }

    /**
     * Lectura de la plantilla.
     * @param template nombre de la plantilla.
     * @return objeto de la liberia que maneja la plantilla.
     */
    public WordprocessingMLPackage readTemplate(String template) {

        try {
            File file = new ClassPathResource(template).getFile();
            log.info("Reading template: " + file.getAbsolutePath());

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);
            wordMLPackage.setTitle("My title");

            return wordMLPackage;

        } catch (Docx4JException e) {
            log.info("Error reading template: " + e.getMessage());
        } catch (Exception e) {
            log.info("File not found: " + template);
        }
        return null;
    }

    /**
     * Metodo que permite eliminar una tabla del documento buscandola por su identifcador.
     * El identificador de tabla unico que se define en:
     * Table properties-> Pestaña "Alt text" y campo "Title". 
     * 
     * @param tabla tabla a eliminar
     */
    public static void removeTableFromDocument(Tbl tabla) {
    	
    	if (tabla != null) {
    		log.info("Eliminando tabla: '{}'", getTableName(tabla));
    		Body parent = (Body) tabla.getParent();    		
    		parent.getContent().remove(tabla);
    		log.info("Tabla: '{}' Eliminada", getTableName(tabla));
    	}   
    	
    }
    
    /**
     * Metodo que encuentra el objeto tbl en el documento para la tabla de identificador id.
     * 
     * @param mainDocumentPart Documento.
     * @param id String identificador de la tabla
     * @return null si no se encuentra o Tbl si se encuentra la tabla.
     */
    public static Tbl findTableById(MainDocumentPart mainDocumentPart, String id) {
    	
    	log.info("Dentro findTableById buscando '{}'", id);    	
    	TableFinder finder = new TableFinder();
        finder.walkJAXBElements(mainDocumentPart.getContent());
        
        for (Object o : finder.tblList) {
        	
        	Object o2 = XmlUtils.unwrap(o); 
        	
        	if (o2 instanceof org.docx4j.wml.Tbl) {
                Tbl tbl = (Tbl)o2;                
                if (isThisTable(id, tbl)){
                	log.info("Table found: '{}'", id);
                	return tbl;
                }  
        	}
        }        
        return null;
    }
    
    /**
     * Metodo que indica si la tabla es la que se busca por identificador
     * @param id identificador 
     * @param tbl tabla
     * @return true si el identificador es el Alt Text ->Title de la tabla.
     */
    public static boolean isThisTable(String id, Tbl tbl) {
    	return tbl.getTblPr().getTblCaption() != null && id.equals(getTableName(tbl));
    }
    

    /**
     * Metodo que clona la fila 1 (Revisar si no tiene cabecera) y sustituye las variables haciendolas dinamicas.
     * 
     * @param tabla Tabla
     * @param numLinesAdd numero de lineas a añadir y sustituir sus variables.
     */
    public static void generateLines(Tbl tabla, int numLinesAdd) {
    	log.info("generateLines para tabla: '{}'. Numero de lineas a crear: {}", getTableName(tabla), numLinesAdd);
    	List<Object> rows = tabla.getContent();
    	
    	if (rows.size() > 1) {    		
    	    Tr templateRow = (Tr) rows.get(1);    	        	    
	    	for (int i = 0; i < numLinesAdd; i++) {
	    		Tr workingRow = XmlUtils.deepCopy(templateRow);
	    		replacement(workingRow, i);
	    		rows.add(workingRow); 
	    	}
    	    tabla.getContent().remove(templateRow);
    	}
    }
    
    private static String getTableName(Tbl tabla) {
    	return tabla.getTblPr().getTblCaption().getVal();
    }
    
    /**
     * Metodo que sustituye la variable por la variable correlativa.
     * @param workingRow fila de trabajo.
     * @param line linea.
     */
    private static void replacement(Tr workingRow, int line) {    	
    	List<Object>  l = getAllElementFromObject(workingRow, Text.class);
    	
    	log.info("replacement. Cambiando la fila: {}", line);
    	
    	for (Object t : l) {    		
    		Text campo = (Text) t;    		
    		String valor = campo.getValue();
    		String nuevoValor = valor.replace("}", "_"+line +"}");
    		log.info("Valor actual: '{}' Nuevo Valor: '{}'", valor, nuevoValor);
    		campo.setValue(nuevoValor);    		
    	}
    }
    
    /**
     * Metodo auxiliar que permite encontrar a partir de un elemento del Documento todos sus hijos (recursion) de un tipo concreto.
     * @param obj objeto origen.
     * @param toSearch Clase de objeto a buscar.
     * @return Listado de Objetos hijos que son de la clase toSearch.
     */
    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }
   
    /**
     * Metodo que graba en Disco el documento Word.
     * @param wordMLPackage documento.
     * @param destinationUrl url donde se grabará.
     */
    public void saveToDisk(WordprocessingMLPackage wordMLPackage, String destinationUrl) {

        try {

            File file = new File(pdfPathUrl);
            FileOutputStream fos = new FileOutputStream(file);

            log.info("Saving to PDF: " + file.getAbsolutePath());

            wordMLPackage.save(fos);

            log.info("File saved to: " + destinationUrl);
        } catch (FileNotFoundException e) {
            log.info("File not found: " + destinationUrl);
        } catch (Docx4JException e) {
            log.info("Error saving file: " + e.getMessage());
        }

    }
}

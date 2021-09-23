package sistema;


import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Comandos {
	    
	    public void digitaComando( List<Produto> produtos ) throws Exception,AWTException{
	        
	       Robot robot = new Robot();
	       
	       produtos.forEach(produto -> { 
	           try {
	        	   robot.keyPress(KeyEvent.VK_ENTER);
	        	   type(produto.getEan());
	        	   robot.keyPress(KeyEvent.VK_ENTER);
	        	   TimeUnit.SECONDS.sleep(1);
	        	   robot.keyPress(KeyEvent.VK_DOWN);
	           } catch ( Exception ex) {
	               Logger.getLogger(Comandos.class.getName()).log(Level.SEVERE, null, ex);
	           }
	       
	       
	       } );
	    }
	    
	    private void type(String s) throws AWTException{
	    
	        Robot robot = new Robot();
	        byte[] bytes = s.getBytes();
	        for (byte b : bytes)
	        {
	            int code = b;
	            // keycode only handles [A-Z] (which is ASCII decimal [65-90])
	            if (code > 96 && code < 123) code = code - 32;
	            robot.delay(80);
	            robot.keyPress(code);
	            robot.keyRelease(code);
	        }
	        robot.keyPress(KeyEvent.VK_ENTER);
	        
	    }
	    
	    private void type(int i) throws AWTException{
	        Robot robot = new Robot();
	        robot.delay(80);
	        robot.keyPress(i);
	        robot.keyRelease(i);
	        robot.keyPress(KeyEvent.VK_ENTER);
	  }
	    
	    
	    
	    
	}

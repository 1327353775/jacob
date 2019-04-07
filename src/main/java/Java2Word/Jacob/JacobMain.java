package Java2Word.Jacob;

/**
 * 测试JacobOperateDoc类
 * @author MrChen
 *
 */
public class JacobMain {

    /**
     * @param args
     */
    public static void main(String[] args) {
        System.out.println("start!");
        System.out.println(System.getProperty("java.library.path")); 
        JacobTest jac = new JacobTest();
        jac.openWord(true);
        jac.createNewDocument();
        jac.insertText("helloworld!");
        jac.saveFileAs("D:\\hell.doc");
        try{
            jac.closeDocument(null);
        }catch(NullPointerException e){
            //捕捉空指针异常，什么也不做
        }
        jac.closeWord();
        System.out.println("end!");
    }

}
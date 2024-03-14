package com.loong.ai.test;

import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import javax.xml.bind.JAXBElement;
import java.util.ArrayList;
import java.util.List;

public class Docx4jHelper {

    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof JaxbXmlPart) {
            List<Object> children = ((JaxbXmlPart<?>) obj).getContent();
            TraversalUtil.visitRecursively(new TraversalUtil.CallbackImpl() {
                public List<Object> apply(Object o) {
                    if (o.getClass().equals(toSearch)) {
                        result.add(o);
                    }
                    return null;
                }
            }, children);
        }
        return result;
    }
}

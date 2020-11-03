/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.microprograms.poi_template;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.Version;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.microprograms.poi_template.config.Configure;
import com.github.microprograms.poi_template.exception.ResolverException;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.render.DefaultRender;
import com.github.microprograms.poi_template.render.Render;
import com.github.microprograms.poi_template.resolver.Resolver;
import com.github.microprograms.poi_template.resolver.TemplateResolver;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.util.PoitlIOUtils;
import com.github.microprograms.poi_template.util.Preconditions;
import com.github.microprograms.poi_template.xwpf.NiceXWPFDocument;

/**
 * The template of word(docx)
 */
public class XWPFTemplate implements Closeable {
    private static Logger logger = LoggerFactory.getLogger(XWPFTemplate.class);
    private static final String SUPPORT_MINIMUM_VERSION = "4.1.1";

    private NiceXWPFDocument doc;

    private Configure config;
    private Resolver resolver;
    private Render renderer;

    private List<MetaTemplate> eleTemplates;

    static {
        try {
            Class.forName("org.apache.poi.Version");
            Preconditions.checkMinimumVersion(Version.getVersion(), SUPPORT_MINIMUM_VERSION,
                    (cur, min) -> "Require Apache POI version at least " + min + ", but now is " + cur
                    + ", please check the dependency of project.");
        } catch (ClassNotFoundException e) {
            // no-op
        }
    }

    private XWPFTemplate() {}

    public static XWPFTemplate compile(String filePath) {
        return compile(new File(filePath));
    }

    public static XWPFTemplate compile(File file) {
        return compile(file, Configure.createDefault());
    }

    public static XWPFTemplate compile(InputStream inputStream) {
        return compile(inputStream, Configure.createDefault());
    }

    public static XWPFTemplate compile(String filePath, Configure config) {
        return compile(new File(filePath), config);
    }

    public static XWPFTemplate compile(File file, Configure config) {
        try {
            return compile(new FileInputStream(file), config);
        } catch (FileNotFoundException e) {
            throw new ResolverException("Cannot find the file [" + file.getPath() + "]", e);
        }
    }

    /**
     * template file as InputStream
     * 
     * @param inputStream
     * @param config
     * @return
     * @version 1.2.0
     */
    public static XWPFTemplate compile(InputStream inputStream, Configure config) {
        try {
            XWPFTemplate template = new XWPFTemplate();
            template.config = config;
            template.doc = new NiceXWPFDocument(inputStream);
            template.resolver = new TemplateResolver(template.config);
            template.renderer = new DefaultRender();
            template.eleTemplates = template.resolver.resolveDocument(template.doc);
            return template;
        } catch (OLE2NotOfficeXmlFileException e) {
            logger.error("Poi-tl currently only supports .docx format");
            throw new ResolverException("Compile template failed", e);
        } catch (IOException e) {
            throw new ResolverException("Compile template failed", e);
        }
    }

    /**
     * Render the template by data model
     * 
     * @param model
     * @return
     */
    public XWPFTemplate render(Object model) {
        this.renderer.render(this, model);
        return this;
    }

    /**
     * Render the template by data model and write to OutputStream, do'not
     * forget invoke #{@link XWPFTemplate#close()},
     * #{@link OutputStream#close()}
     * 
     * @param model
     * @param out
     * @return
     * @throws IOException
     */
    public XWPFTemplate render(Object model, OutputStream out) throws IOException {
        this.render(model);
        this.write(out);
        return this;
    }

    /**
     * bind render policy
     * 
     * @param tagName
     * @param policy
     * @return
     */
    public XWPFTemplate bind(String tagName, RenderPolicy policy) {
        this.config.customPolicy(tagName, policy);
        return this;
    }

    /**
     * write to output stream, do'not forget invoke
     * #{@link XWPFTemplate#close()}, #{@link OutputStream#close()} finally
     * 
     * @param out
     *            eg.ServletOutputStream
     * @throws IOException
     */
    public void write(OutputStream out) throws IOException {
        this.doc.write(out);
    }

    /**
     * write to file
     * 
     * @param path
     * @throws IOException
     */
    public void writeToFile(String path) throws IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(path);
            this.write(out);
            out.flush();
        }
        finally {
            PoitlIOUtils.closeQuietlyMulti(this.doc, out);
        }
    }

    /**
     * reload the template
     * 
     * @param doc
     */
    public void reload(NiceXWPFDocument doc) {
        PoitlIOUtils.closeLoggerQuietly(this.doc);
        this.doc = doc;
        this.eleTemplates = this.resolver.resolveDocument(doc);
    }

    /**
     * close the document
     * 
     * @throws IOException
     */
    @Override
    public void close() throws IOException {
        this.doc.close();
    }

    public List<MetaTemplate> getElementTemplates() {
        return eleTemplates;
    }

    public NiceXWPFDocument getXWPFDocument() {
        return this.doc;
    }

    public Configure getConfig() {
        return config;
    }

    public Resolver getResolver() {
        return resolver;
    }

}

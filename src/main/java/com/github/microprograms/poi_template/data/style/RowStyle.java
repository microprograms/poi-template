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
package com.github.microprograms.poi_template.data.style;

import java.io.Serializable;

import com.github.microprograms.poi_template.util.UnitUtils;

public class RowStyle implements Serializable {

    private static final long serialVersionUID = 1L;
    /**
     * in twips
     * 
     * @see #{@link UnitUtils#cm2Twips()}
     */
    private int height;

    /**
     * default cell style for all cells in the current row
     */
    private CellStyle defaultCellStyle;

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    public void setDefaultCellStyle(CellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }

}

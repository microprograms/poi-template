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
package com.github.microprograms.poi_template.util;

import java.util.Collections;
import java.util.List;

public final class UnitUtils {

    /**
     * cm to twips
     * 
     * @param cm
     * @return in twentieths of a point (1/1440 of an inch)
     */
    public static int cm2Twips(double cm) {
        return (int) (cm / 2.54 * 1440);
    }

    /**
     * point to twips
     * 
     * @param pt
     * @return in twentieths of a point (1/1440 of an inch)
     */
    public static int point2Twips(int pt) {
        return pt * 20;
    }

    /**
     * average the width
     * 
     * @param width
     * @param col
     * @return
     */
    public static int[] average(int width, int col) {
        int colVal = (Integer.valueOf(width)) / col;
        List<Integer> nCopies = Collections.nCopies(col, colVal);
        return nCopies.stream().mapToInt(i -> i).toArray();
    }

}

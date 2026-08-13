package com.office.automation;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.HashMap;
import java.util.Map;

/**
 * Utility class for loading data from various sources
 */
public class DataLoader {
    private static final Logger logger = LoggerFactory.getLogger(DataLoader.class);
    private static final Gson gson = new Gson();

    /**
     * Loads data from a JSON file
     */
    public static Map<String, String> loadFromJson(String filePath) throws IOException {
        try (FileReader reader = new FileReader(filePath)) {
            Type type = new TypeToken<Map<String, String>>(){}.getType();
            Map<String, String> data = gson.fromJson(reader, type);
            logger.info("Loaded {} entries from JSON file", data.size());
            return data != null ? data : new HashMap<>();
        }
    }

    /**
     * Creates a simple data map (useful for direct API usage)
     */
    public static Map<String, String> createDataMap(Object... pairs) {
        if (pairs.length % 2 != 0) {
            throw new IllegalArgumentException("Pairs must have even number of elements");
        }

        Map<String, String> data = new HashMap<>();
        for (int i = 0; i < pairs.length; i += 2) {
            data.put(pairs[i].toString(), pairs[i + 1].toString());
        }
        return data;
    }
}

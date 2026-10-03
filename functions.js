"use strict";
(() => {
  // src/functions/json-to-entity.ts
  function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
  }
  function isEntityCellValue(value) {
    return isRecord(value) && value.type === "Entity" && typeof value.text === "string" && isRecord(value.properties);
  }
  function isArrayEntityValue(value) {
    var _a;
    const propertyNames = Object.keys(value.properties);
    return propertyNames.length === 1 && propertyNames[0] === "items" && ((_a = value.properties.items) == null ? void 0 : _a.type) === "Array";
  }
  function setOwnValue(target, key, value) {
    Object.defineProperty(target, key, {
      value,
      enumerable: true,
      configurable: true,
      writable: true
    });
  }
  function createEmptyCellValue() {
    return { type: "Empty" };
  }
  function parseJsonValue(value) {
    if (value === null || typeof value === "string" || typeof value === "boolean") {
      return value;
    }
    if (typeof value === "number") {
      if (!Number.isFinite(value)) {
        throw new TypeError("JSON \u306E\u6570\u5024\u306F\u6709\u9650\u5024\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
      }
      return value;
    }
    if (Array.isArray(value)) {
      return value.map(parseJsonValue);
    }
    if (isRecord(value)) {
      const parsedObject = {};
      for (const [key, entry] of Object.entries(value)) {
        setOwnValue(parsedObject, key, parseJsonValue(entry));
      }
      return parsedObject;
    }
    throw new TypeError("JSON \u3068\u3057\u3066\u6271\u3048\u306A\u3044\u5024\u304C\u542B\u307E\u308C\u3066\u3044\u307E\u3059\u3002");
  }
  function toCellValue(value, keyForNestedEntity) {
    if (value === null) {
      return createEmptyCellValue();
    }
    switch (typeof value) {
      case "string":
        return { type: "String", basicType: "String", basicValue: value };
      case "number":
        return { type: "Double", basicType: "Double", basicValue: value };
      case "boolean":
        return { type: "Boolean", basicType: "Boolean", basicValue: value };
      case "object":
        return Array.isArray(value) ? convertArray(value, keyForNestedEntity) : convertObject(value, keyForNestedEntity);
    }
  }
  function toSpillCellValue(value, keyForNestedEntity) {
    if (value === null) {
      return createEmptyCellValue();
    }
    switch (typeof value) {
      case "string":
      case "number":
      case "boolean":
        return value;
      case "object":
        return Array.isArray(value) ? convertArray(value, keyForNestedEntity) : convertObject(value, keyForNestedEntity);
    }
  }
  function convertArray(values, text) {
    const elements = values.map(
      (value, index) => Array.isArray(value) ? value.map((item) => toCellValue(item, String(index))) : [toCellValue(value, String(index))]
    );
    return {
      type: "Entity",
      text,
      properties: {
        items: { type: "Array", elements }
      }
    };
  }
  function convertObject(value, text) {
    const properties = {};
    for (const [key, entry] of Object.entries(value)) {
      setOwnValue(properties, key, toCellValue(entry, key));
    }
    return { type: "Entity", text, properties };
  }
  function parseJsonText(jsonText) {
    if (jsonText.trim() === "") {
      throw new TypeError("JSON \u6587\u5B57\u5217\u304C\u7A7A\u3067\u3059\u3002");
    }
    const parsed = JSON.parse(jsonText);
    return parseJsonValue(parsed);
  }
  function getObjectEntrySource(input) {
    if (typeof input === "string") {
      const value = parseJsonText(input);
      if (!isRecord(value)) {
        throw new TypeError("JSON \u306E\u6700\u4E0A\u4F4D\u306F\u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
      }
      const entries2 = Object.entries(value);
      if (entries2.length === 0) {
        throw new TypeError("\u7A7A\u306E JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306F Excel \u306Bspill\u3067\u304D\u307E\u305B\u3093\u3002");
      }
      return { kind: "json", entries: entries2 };
    }
    if (!isEntityCellValue(input) || isArrayEntityValue(input)) {
      throw new TypeError("JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u307E\u305F\u306F Excel Entity \u3092\u6307\u5B9A\u3057\u3066\u304F\u3060\u3055\u3044\u3002");
    }
    const entries = Object.entries(input.properties);
    if (entries.length === 0) {
      throw new TypeError("\u7A7A\u306E JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306F Excel \u306Bspill\u3067\u304D\u307E\u305B\u3093\u3002");
    }
    return { kind: "entity", entries };
  }
  function toSpillCellValueFromEntity(value) {
    switch (value.type) {
      case "Empty":
        return createEmptyCellValue();
      case "String":
      case "Double":
      case "Boolean":
        return value.basicValue;
      case "Entity":
        return value;
      case "Array":
        throw new TypeError("Excel Array \u306F\u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306E\u30D7\u30ED\u30D1\u30C6\u30A3\u5024\u3068\u3057\u3066spill\u3067\u304D\u307E\u305B\u3093\u3002");
    }
  }
  function parseJsonObjectKeys(input) {
    return getObjectEntrySource(input).entries.map(([key]) => [key]);
  }
  function parseJsonObjectValues(input) {
    const source = getObjectEntrySource(input);
    if (source.kind === "json") {
      return source.entries.map(([key, value]) => [toSpillCellValue(parseJsonValue(value), key)]);
    }
    return source.entries.map(([, value]) => [toSpillCellValueFromEntity(value)]);
  }
  function isEmptyJsonObject(input) {
    if (typeof input === "string") {
      const value = parseJsonText(input);
      return isRecord(value) && Object.keys(value).length === 0;
    }
    if (isRecord(input) && input.type === "Entity") {
      const properties = input.properties;
      if (properties === void 0) {
        return true;
      }
      if (!isRecord(properties)) {
        throw new TypeError("Excel Entity \u306E properties \u304C\u4E0D\u6B63\u3067\u3059\u3002");
      }
      return Object.keys(properties).length === 0;
    }
    throw new TypeError("JSON \u6587\u5B57\u5217\u307E\u305F\u306F Excel Entity \u3092\u6307\u5B9A\u3057\u3066\u304F\u3060\u3055\u3044\u3002");
  }
  function convertToEntity(value, text) {
    if (Array.isArray(value)) {
      return convertArray(value, text);
    }
    if (isRecord(value)) {
      return convertObject(value, text);
    }
    throw new TypeError("JSON \u306E\u6700\u4E0A\u4F4D\u306F\u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u307E\u305F\u306F\u914D\u5217\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
  }
  function parseJson(jsonText) {
    const value = parseJsonText(jsonText);
    if (Array.isArray(value) || isRecord(value)) {
      return convertToEntity(value, "JSON Entity");
    }
    return value === null ? createEmptyCellValue() : value;
  }
  function parseJsonObject(jsonText) {
    const value = parseJsonText(jsonText);
    if (!isRecord(value)) {
      throw new TypeError("JSON \u306E\u6700\u4E0A\u4F4D\u306F\u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
    }
    return convertObject(value, "JSON Entity");
  }
  function parseJsonArray(jsonText) {
    const value = parseJsonText(jsonText);
    if (!Array.isArray(value)) {
      throw new TypeError("JSON \u306E\u6700\u4E0A\u4F4D\u306F\u914D\u5217\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
    }
    if (value.length === 0) {
      throw new TypeError("\u7A7A\u306E JSON \u914D\u5217\u306F Excel \u306B\u30B9\u30D4\u30EB\u3067\u304D\u307E\u305B\u3093\u3002");
    }
    const rows = value.map(
      (entry, index) => Array.isArray(entry) ? entry.map((item) => toSpillCellValue(item, String(index))) : [toSpillCellValue(entry, String(index))]
    );
    const columnCount = rows[0].length;
    if (columnCount === 0 || rows.some((row) => row.length !== columnCount)) {
      throw new TypeError("JSON \u914D\u5217\u306F\u540C\u3058\u5217\u6570\u306E\u884C\u3067\u69CB\u6210\u3055\u308C\u3066\u3044\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059\u3002");
    }
    return rows;
  }

  // src/functions/functions.ts
  function json(jsonText) {
    try {
      return parseJson(jsonText);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u3092 Excel \u306E\u5024\u306B\u5909\u63DB\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("ENTITY", json);
  function object(jsonText) {
    try {
      return parseJsonObject(jsonText);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306B\u5909\u63DB\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("OBJECT", object);
  function objectKeys(input) {
    try {
      return parseJsonObjectKeys(input);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306E\u30AD\u30FC\u3092spill\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("OBJECTKEYS", objectKeys);
  function objectValues(input) {
    try {
      return parseJsonObjectValues(input);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u306E\u5024\u3092spill\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("OBJECTVALUES", objectValues);
  function isEmptyObject(input) {
    try {
      return isEmptyJsonObject(input);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u30AA\u30D6\u30B8\u30A7\u30AF\u30C8\u3092\u5224\u5B9A\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("ISEMPTYOBJECT", isEmptyObject);
  function array(jsonText) {
    try {
      return parseJsonArray(jsonText);
    } catch (error) {
      const message = error instanceof Error ? error.message : "JSON \u914D\u5217\u3092 Excel \u306B\u30B9\u30D4\u30EB\u3067\u304D\u307E\u305B\u3093\u3002";
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
  }
  CustomFunctions.associate("ARRAY", array);
})();

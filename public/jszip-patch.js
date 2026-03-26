(function patchJSZipV2Compat() {
  if (typeof JSZip === "undefined") return;
  var _origLoad = JSZip.loadAsync.bind(JSZip);
  JSZip.loadAsync = function (data, opts) {
    return _origLoad(data, opts).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) {
        return !zip.files[n].dir;
      });

      // Pre-fetch every file so we can expose sync-like methods
      return Promise.all(names.map(function (name) {
        return zip.files[name].async("arraybuffer").then(function (ab) {
          var entry = zip.files[name];
          entry.asArrayBuffer = function () {
            // Return a fresh slice to avoid contamination
            return ab.slice(0);
          };
          entry.asUint8Array = function () { return new Uint8Array(ab); };
          entry.asBinary = function () {
            var s = "";
            var u8 = new Uint8Array(ab);
            for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
            return s;
          };
          entry.asText = function () {
            return new TextDecoder("utf-8").decode(new Uint8Array(ab));
          };
        });
      })).then(function () {
        // Override zip.file(name) to be fuzzy
        var _origFile = zip.file.bind(zip);
        zip.file = function(name) {
          if (!name) return null;
          // 1. Exact match
          var exact = _origFile(name);
          if (exact) return exact;

          // 2. Fuzzy match (ignore case and leading slashes)
          var cleanName = name.toLowerCase().replace(/\\/g, '/');
          try { cleanName = decodeURIComponent(cleanName); } catch(e) {}
          if (cleanName.startsWith('/')) cleanName = cleanName.substring(1);
          if (cleanName.startsWith('ppt/')) cleanName = cleanName.substring(4); // Remove ppt/ to match just media/...
          if (cleanName.startsWith('../')) cleanName = cleanName.substring(3);

          var keys = Object.keys(zip.files);
          for (var i = 0; i < keys.length; i++) {
            var lowKey = keys[i].toLowerCase();
            try { lowKey = decodeURIComponent(lowKey); } catch(e) {}
            if (lowKey.endsWith(cleanName)) {
              return zip.files[keys[i]];
            }
          }

          // 3. Fallback dummy to prevent PPTX engine crash
          console.warn("Fuzzy match failed for image/resource:", name);
          return {
            asArrayBuffer: function() { return new ArrayBuffer(0); },
            asUint8Array: function() { return new Uint8Array(0); },
            asText: function() { return ""; },
            asBinary: function() { return ""; }
          };
        };

        return zip;
      });
    });
  };
})();

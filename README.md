# ClickHere Interop Services Workgroup Machines

A VB.NET console program that lists the computers in a Windows
workgroup.

## What it does

It calls the Win32 `NetServerEnum` function from `Netapi32.dll`
through interop services and prints the machines found in the
`CLICKHERE` workgroup.

## Tech

- Visual Basic .NET console application
- Win32 `Netapi32` API

## Building

Open the `.sln` file in Visual Studio and build it. Windows only.

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).

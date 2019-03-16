# VB6 and .NET Standard Class Library Interop
This sample shows how to call methods in a .NET Standard 2.0 class library from a Visual Basic 6 application. You can also use Full .NET Framework (4.7.2 or other versions) for your class libraries and it will make no difference.

### Step 1 - Create .NET Standard Class Library

* Create an interface for the methods you need to use from VB6 and make it visible to COM

```csharp
[Guid("45f838c9-d8cd-4198-b2b3-f4e8c2a5b588")]
[ComVisible(true)]
public interface IMathUtils
{
    double CalculateComplexMethod();
    void CommandWithParameters(int number1, int number2);
}
```

> Guid must be unique for this interface

* Implement the interface in a class and also make it visible to COM. Don't use the Guid as the interface

```csharp
[Guid("293bcd3a-e771-45d5-8a53-5413997c2de8")]
[ComVisible(true)]
[ProgId("NetStandardLibrary.MathUtils")]
public class MathUtils : IMathUtils
{
    // For COM interop we need parameterless constructor
    public MathUtils() { }

    public double CalculateComplexMethod() => 2 * 2;

    public void CommandWithParameters(int number1, int number2)
    {
        // Do something in here like saving to a JSON file, POST to some REST API or anything
        // without needing to code in prehistoric VB6
    }
}
```


### Extras
##### VB6 in Github

In order to use Github for your VB6 source code add this extensions to your Visual Studio .gitignore file. Like in this project [.gitignore file for VB6](.gitignore)

```markdown
*.csi
*.exp
*.lib
*.lvw
*.dca
```

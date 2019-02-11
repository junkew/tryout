#Run as admin
enable-psremoting

[DSCLocalConfigurationManager()]
configuration LCMConfig
{
    Node localhost
    {
        Settings
        {
            RefreshMode = 'Push'
        }
    }
}

LCMConfig -OutputPath LCMConfig

Set-DscLocalConfigurationManager LCMConfig



configuration HelloWorld {
 param ()
 Import-DscResource –ModuleName 'PSDesiredStateConfiguration'
 
 Node Localhost 
 {
   Environment EnvironmentExample
   {
        Ensure = "Present"  # You can also set Ensure to "Absent"
        Name = "TestEnvironmentVariable"
        Value = "TestValue you should see used in the script part"
   }

    Script ScriptExample
    {
        SetScript = {
            [System.Environment]::SetEnvironmentVariable("FOOBAR", "hi there")
            $sw = New-Object System.IO.StreamWriter("c:\TestFile.txt")
            $sw.WriteLine("Some sample string")
            $sw.WriteLine([System.Environment]::GetEnvironmentVariable("TestEnvironmentVariable") )
            $sw.WriteLine([System.Environment]::GetEnvironmentVariable("FOOBAR") )
            $sw.Close()
        }
        TestScript = { Test-Path "c:\TestFile.txt" }
        GetScript = { @{ Result = (Get-Content c:\TestFile.txt) } }
        DependsOn="[Environment]EnvironmentExample"
    }

   # Create a Test File
   File CreateTestFile
   {
     Ensure          = "Present"
     DestinationPath = "C:\example.txt"
     Contents        = "Hello World."
     Type            = "File"
   }
    # Create a Test File
   File CreateTestFile2
   {
     Ensure          = "Present"
     DestinationPath = "C:\example2.txt"
     Contents        = "Hello World 2."
     Type            = "File"
   }
 }

}
 
# Create MOF Files
HelloWorld -OutputPath HelloWorld

configuration HelloWorld2 {
 param ()
  Import-DscResource –ModuleName 'PSDesiredStateConfiguration'
 Node Localhost 
 {
   # Create a Test File
   File CreateTestFile
   {
     Ensure          = "Present"
     DestinationPath = "C:\example3.txt"
     Contents        = "Hello World 3."
     Type            = "File"
   }
 }
}
# Create MOF Files
HelloWorld2 -OutputPath HelloWorld2
 
# Start DSC Configuration
#Start-DscConfiguration -Path C:\ScriptimusExMachina\HelloWorld -ComputerName Localhost -Verbose -Wait
Start-DscConfiguration -Path HelloWorld -ComputerName Localhost -Verbose -Wait
#Start-DscConfiguration -Path HelloWorld2 -ComputerName Localhost -Verbose -Wait

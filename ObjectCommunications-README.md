# Object-Communications-Global-Variable-Alternative

### A general message passing mechanism used to pass messages (data) between Objects. Does not require both Objects to be active. Uses only one Global variable as an Instance Variable for the Class. Alleviates the need for dangerous Global/Public variables and replaces OpenArgs. Simplifies all Object communications because all such communication can be done the same way.

This Class implements messaging between any two objects. It allows two or more Objects to exchange messages, or data, in a secure manner because no Public variables are used except for the Instance Variable. 

One object inserts one or more Values as Key/Value pairs and another object retrieves the Values by their Keys. The Objects only need to agree on the names of the Keys used to exchange the data, besides the name of the Instance Variable. 

This Class eliminates the need to create unique Class Properties in every Form or Report that must communicate with another Object. In addition, the Form or Report need not be active in order for the receiver to access the data as long as the Instance Variable is Public.

The same Instance Variable can be used for an entire program because, in addition to the Key, there is an optional Node that may be assigned to each Key. Nodes are qualifiers or prefixes to Keys. So, if there is ever any doubt about, or possibility of, naming conflicts, just declare and use a Node. In reality, there is a default Node that is always in effect if none is explicitly specified.

The same effect could be accomplished by creating Properties in one Object and then invoking those Properties from the other Object but creating all of these Properties each time for each Object has gotten tedious. Other solutions that are used are hidden forms and Globals. Hidden forms work but are too labor-intensive, create too much memory overhead, and are inelegant. Globals are universally considered to be a poor solution to the problem for some very well-known and accepted reasons.

OpenArgs can be used to pass data into an Object, and that is a good way to do it but you can't use it to pass data back to the invoking Object so OpenArgs is not a complete, universal solution to the problem.

This Class also supports asychronous communications if a Public Instance Variable is used because each Object actually communicates with this Class which internally stores the passed data. The receiver can read the data from the Class whenever it wants the data. The sender does not have to be active. 

There are times when a Public variable is necessary. This Class can be used to create and protect Public variables from name collisions and inadvertant over-writing of a value. If there is any doubt if a Key has already been used, the it specifying a unique Node avoid any accidental trashing of Public variables. Admittedly, this is not a novel solution as others have proposed protecting Public variable by making them Static variables inside a Module and writing a functions to read and write each variable. This is just another use of this Class that one may want to take advantage of.

This Class uses the Collection object rather than the Scripting.Dictionary object because the Collection object does not require a Reference. Since the complexity of emulating the Key/Value structure of Scripting.Dictionary using the simpler Collection object is hidden  within the Class, it was decided that trading this extra coding complexity to reduce and avoid installation and configuration difficulties of the application was worth the coding effort. The API would be the same for either implementation.

USAGE:
        1. Dim/Public col as [New] Comm_cls (use New for autovivification)
        2. col.AddValue(Key As String, Value As Variant, Optional NodeName As String = "$Default")
        3) col.GetValue(Key As String, Optional NodeName As String = "$Default")
        4) col.DeleteItem(Key As String, Optional Node As String = "$Default")
        5) col.DeleteKey(Node As String)
        
There is no Create Method because the AddValue method automatically creates the Key if it doesn't exist before storing the Value. 

There is no Update Method because the AddValue method will update the Value if the Key exists.

REQUIRES: Error Message Handler: ErrHandler in module ErrorMessageHandler_Lib

Please send any questions, comments, or bug reports to Paul@PStrauss.net

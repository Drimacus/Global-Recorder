Global-Recorder
===============

Record events outside of Excel.

In order to use:

Need to add reference to Microsoft Visual Basic for Applications Extensibility 5.3

Need to set macro security options : Trust access to the VBA project model

Run the macro 'startRecording' to start logging mouse and keyboard events in a new module.

If stopped with stop button, add the line 'EndSub' at the end of the recorded code. Otherwise assign a control to the macro 'stopRecording' to have this done for you.

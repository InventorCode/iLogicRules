'Show Front on View Cube'
ThisApplication.CommandManager.ControlDefinitions.Item("AppFrontViewCmd").Execute




'''''''''''''''
'set current as front view
ThisApplication.CommandManager.ControlDefinitions("AppViewCubeViewFrontCmd").Execute

'set to iso view
Dim oCamera As Camera 
oCamera = ThisApplication.ActiveView.Camera 
oCamera.ViewOrientationType = 10760 'Iso Top Left View Orientation 
oCamera.Apply
'set current iso as home view
ThisApplication.CommandManager.ControlDefinitions("AppViewCubeViewHomeFloatingCmd").Execute

'return to front
oCamera.ViewOrientationType = 10764 'Front View Orientation 
oCamera.Apply 

''''''''''''
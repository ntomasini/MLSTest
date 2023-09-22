
Public Structure Species
    Dim length As Double   'distance value
    Dim spName As String    'species name
    Dim nodes() As Integer  'list of nodes that this sp appears in
    Dim y As Double        'the y position of RIGHT HAND END of sp branch
    Dim x As Double        'the x position of RIGHT HAND END of sp branch
End Structure


Public Structure Node
    Dim length As Double     'distance value
    Dim bootstrap As Double  'bootstrap value
    Dim nodename As String
    Dim x_left As Double     'x position of left hand end of node distance
    Dim x_right As Double    'x position of right hand end of node distance
    '(the end of this node's vertical line)
    Dim y_top As Double      'top of this node's vertical line
    Dim y_bottom As Double   'bottom of this node's vertical
    Dim y_mid As Double      'middle of vertical - where this node's distance
    'comes from
    Dim innernodes() As Integer   'indexes of all nodes within this one
    Dim innerUniqueNodes() As Integer  'indexes of all nodes ONLY within this one
    Dim nodeCount As Integer      'inner node count
    Dim uniqueNodeCount As Integer  'inner unique node count
    Dim spCount As Integer        'number of species that come off this node only
    Dim spec() As Integer         'indexes of all sp that come off this node only
End Structure



Module Moduletreeviewer

    
End Module

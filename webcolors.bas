Attribute VB_Name = "webcolors"
Private Const mod_name As String = "webcolors"
Private Const module_author As String = "Ben Fisher"
Private Const module_version As String = "0.0.3"

'Module that contains the 140 named webcolors as Long
'simply add this module to your project to be able to access
'the named colors

'Pink colors
Public Const mediumvioletred As Long = 8721863
Public Const deeppink As Long = 9639167
Public Const palevioletred As Long = 9662683
Public Const hotpink As Long = 11823615
Public Const lightpink As Long = 12695295
Public Const pink As Long = 13353215

'Red colors
Public Const darkred As Long = 139
Public Const red As Long = 255
Public Const firebrick As Long = 2237106
Public Const crimson As Long = 3937500
Public Const indianred As Long = 6053069
Public Const lightcoral As Long = 8421616
Public Const salmon As Long = 7504122
Public Const darksalmon As Long = 8034025
Public Const lightsalmon As Long = 8036607

'Orange colors
Public Const orangered As Long = 17919
Public Const tomato As Long = 4678655
Public Const darkorange As Long = 36095
Public Const coral As Long = 5275647
Public Const orange As Long = 42495

'Yellow colors
Public Const darkkhaki As Long = 7059389
Public Const gold As Long = 55295
Public Const khaki As Long = 9234160
Public Const peachpuff As Long = 12180223
Public Const yellow As Long = 65535
Public Const palegoldenrod As Long = 11200750
Public Const moccasin As Long = 11920639
Public Const papayawhip As Long = 14020607
Public Const lightgoldenrodyellow As Long = 13826810
Public Const lemonchiffon As Long = 13499135
Public Const lightyellow As Long = 14745599

'Brown colors
Public Const maroon As Long = 128
Public Const brown As Long = 2763429
Public Const saddlebrown As Long = 1262987
Public Const sienna As Long = 2970272
Public Const chocolate As Long = 1993170
Public Const darkgoldenrod As Long = 755384
Public Const peru As Long = 4163021
Public Const rosybrown As Long = 9408444
Public Const goldenrod As Long = 2139610
Public Const sandybrown As Long = 6333684
Public Const tan As Long = 9221330
Public Const burlywood As Long = 8894686
Public Const wheat As Long = 11788021
Public Const navajowhite As Long = 11394815
Public Const bisque As Long = 12903679
Public Const blanchedalmond As Long = 13495295
Public Const cornsilk As Long = 14481663

'Purple and magenta colors
Public Const indigo As Long = 8519755
Public Const purple As Long = 8388736
Public Const darkmagenta As Long = 9109643
Public Const darkviolet As Long = 13828244
Public Const darkslateblue As Long = 9125192
Public Const blueviolet As Long = 14822282
Public Const darkorchid As Long = 13382297
Public Const fuchsia As Long = 16711935
Public Const magenta As Long = 16711935
Public Const slateblue As Long = 13458026
Public Const mediumslateblue As Long = 15624315
Public Const mediumorchid As Long = 13850042
Public Const mediumpurple As Long = 14381203
Public Const orchid As Long = 14053594
Public Const violet As Long = 15631086
Public Const plum As Long = 14524637
Public Const thistle As Long = 14204888
Public Const lavender As Long = 16443110

'Blue colors
Public Const midnightblue As Long = 7346457
Public Const navy As Long = 8388608
Public Const darkblue As Long = 9109504
Public Const mediumblue As Long = 13434880
Public Const blue As Long = 16711680
Public Const royalblue As Long = 14772545
Public Const steelblue As Long = 11829830
Public Const dodgerblue As Long = 16748574
Public Const deepskyblue As Long = 16760576
Public Const cornflowerblue As Long = 15570276
Public Const skyblue As Long = 15453831
Public Const lightskyblue As Long = 16436871
Public Const lightsteelblue As Long = 14599344
Public Const lightblue As Long = 15128749
Public Const powderblue As Long = 15130800

'Cyan colors
Public Const teal As Long = 8421376
Public Const darkcyan As Long = 9145088
Public Const lightseagreen As Long = 11186720
Public Const cadetblue As Long = 10526303
Public Const darkturquoise As Long = 13749760
Public Const mediumturquoise As Long = 13422920
Public Const turquoise As Long = 13688896
Public Const aqua As Long = 16776960
Public Const cyan As Long = 16776960
Public Const aquamarine As Long = 13959039
Public Const paleturquoise As Long = 15658671
Public Const lightcyan As Long = 16777184

'Green colors
Public Const darkgreen As Long = 25600
Public Const green As Long = 32768
Public Const darkolivegreen As Long = 3107669
Public Const forestgreen As Long = 2263842
Public Const seagreen As Long = 5737262
Public Const olive As Long = 32896
Public Const olivedrab As Long = 2330219
Public Const mediumseagreen As Long = 7451452
Public Const limegreen As Long = 3329330
Public Const lime As Long = 65280
Public Const springgreen As Long = 8388352
Public Const mediumspringgreen As Long = 10156544
Public Const darkseagreen As Long = 9419919
Public Const mediumaquamarine As Long = 11193702
Public Const yellowgreen As Long = 3329434
Public Const lawngreen As Long = 64636
Public Const chartreuse As Long = 65407
Public Const lightgreen As Long = 9498256
Public Const greenyellow As Long = 3145645
Public Const palegreen As Long = 10025880

'White colors
Public Const mistyrose As Long = 14804223
Public Const antiquewhite As Long = 14150650
Public Const linen As Long = 15134970
Public Const beige As Long = 14480885
Public Const whitesmoke As Long = 16119285
Public Const lavenderblush As Long = 16118015
Public Const oldlace As Long = 15136253
Public Const aliceblue As Long = 16775408
Public Const seashell As Long = 15660543
Public Const ghostwhite As Long = 16775416
Public Const honeydew As Long = 15794160
Public Const floralwhite As Long = 15792895
Public Const azure As Long = 16777200
Public Const mintcream As Long = 16449525
Public Const snow As Long = 16448255
Public Const ivory As Long = 15794175
Public Const white As Long = 16777215

'Black and gray colors
Public Const black As Long = 0
Public Const darkslategray As Long = 5197615
Public Const dimgray As Long = 6908265
Public Const slategray As Long = 9470064
Public Const gray As Long = 8421504
Public Const lightslategray As Long = 10061943
Public Const darkgray As Long = 11119017
Public Const silver As Long = 12632256
Public Const lightgray As Long = 13882323
Public Const gainsboro As Long = 14474460

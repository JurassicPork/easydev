Arrays
======

Append
------

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    a = util.append(a, "Sofia")
    util.msgbox( a )

Delete
------

.. code-block:: vbnet

    a = util.delete(a, "Nikole")
    util.msgbox( a )

Extend
------

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    a2 = Array("Sofia", "Anita")
    a = util.extend(a, a2)
    util.msgbox( a )

Multiplicate
------------

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    a = util.multi(a, 2)
    util.msgbox( a )

Unique values
-------------

.. code-block:: vbnet

    a = Array(1,2,"Two",3,3,3,4,4,4,4,5,5,5,5,5,"Uno","Uno")
    a = util.unique(a)
    util.msgbox( a )

Reverse
-------

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    a = util.reverse(a)
    util.msgbox( a )

Insert
------

Insert element in pos

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    a = util.insert(a, 2, "Mary")
    util.msgbox( a )

Remove
------

Remove element in pos and return array and element

.. code-block:: vbnet

    a = Array(1,2,"Two",3,3,3,4,4,4,4,5,5,5,5,5,"Uno","Uno")
    data = util.pop(a, 2)
    util.msgbox( data(0) )  'Array without element in pos
    util.msgbox( data(1) )  'Element removed

Remove first element found

.. code-block:: vbnet

    a = Array(1,2,2,3,3,3,4,4,4,4,5,5,5,5,5,"Uno","Uno")
    util.msgbox( util.remove(a, 5, False) )

Remove all elements found

.. code-block:: vbnet

    util.msgbox( util.remove(a, 5, True) )

Len
---

.. code-block:: vbnet

    a = Array(1,2,2,3,3,3,4,4,4,4,5,5,5,5,5,"Uno","Uno")
    util.msgbox( util.len(a) )

Count
-----

.. code-block:: vbnet

    a = Array(1,2,2,3,3,3,4,4,4,4,5,5,5,5,5,"Uno","Uno")
    util.msgbox( util.count(a, 3) )
    util.msgbox( util.count(a, 5) )
    util.msgbox( util.count(a, "Uno") )

Index
-----

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion")
    util.msgbox( util.index(a, "Naomi") )
    util.msgbox( util.index(a, "Monica") )

Max, Min and Average
--------------------

.. code-block:: vbnet

    a = Array(1,2,3,4,5,6,7,8,9,10)
    util.msgbox( util.max(a) )
    util.msgbox( util.min(a) )
    util.msgbox( util.average(a) )

Sum
---

.. code-block:: vbnet

    a = Array(1,2,3,4,5,6,7,8,9,10)
    util.msgbox( util.sum(a) )

Only sum values, the first element is string

.. code-block:: vbnet

    a = Array("10", 1,2,3,4,5,6,7,8,9,10, "One", "Two")
    util.msgbox( util.sum(a) )

Exists
------

If value exists in array

.. code-block:: vbnet

    a = Array(1,2,3,4,5,"One","Seven",9,10)
    util.msgbox( util.exists(a, "One") )
    util.msgbox( util.exists(a, "Two") )

Equal
-----

If array 1 is equal to array2

.. code-block:: vbnet

    a1 = Array(1,2,3) : a2 = Array(1,2,3)
    util.msgbox( util.equal(a1, a2) )

    a1 = Array(1,"Dos",3) : a2 = Array(1,2,"Tres")
    util.msgbox( util.equal(a1, a2) )


Slice
-----

Copy

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion","Sofia","Anita")
    a2 = util.slice(a, "[:]")
    util.msgbox( a2 )

First two elements

.. code-block:: vbnet

    a2 = util.slice(a, "[:2]")
    util.msgbox( a2 )

Last two elements

.. code-block:: vbnet

    a2 = util.slice(a, "[-2:]")
    util.msgbox( a2 )

Range

.. code-block:: vbnet

    a2 = util.slice(a, "[2:-2]")
    util.msgbox( a2 )

    a2 = util.slice(a, "[::2]")
    util.msgbox( a2 )

    a2 = util.slice(a, "[1::2]")
    util.msgbox( a2 )

Reverse

.. code-block:: vbnet

    a2 = util.slice(a, "[::-1]")
    util.msgbox( a2 )


Sorted
------

Sorted unidimension array

.. code-block:: vbnet

    a = Array("Nikole","Scarlett","Monica","Naomi","Marion","Sofia","Anita")
    a = util.sorted(a, 0)
    util.msgbox( a )

Sorted multidimension array

.. code-block:: vbnet

    a = Array( _
        Array(1, 1, 3, "a", 56), _
        Array(1, 2, 3, "z", 43), _
        Array(1, 3, 3, "g", 78), _
        Array(1, 4, 3, "e", 32), _
        Array(1, 5, 3, "M", 89) _
    )
    a = util.sorted(a, 0)
    util.msgbox( a )
    a = util.sorted(a, 1)
    util.msgbox( a )
    a = util.sorted(a, 2)
    util.msgbox( a )
    a = util.sorted(a, 3)
    util.msgbox( a )
    a = util.sorted(a, 4)
    util.msgbox( a )

Get column

.. code-block:: vbnet

    util.msgbox(util.getColumn(a, 1))


Operations
----------

.. code-block:: vbnet

    Sub ArraysOperations()
        util = createUnoService("org.universolibre.EasyDev")

        a1 = Array(1,2,3,4,5) : a2 = Array(3,4,5,6,7,8)
        a = util.union(a1, a2)
        util.msgbox( a )

        a = util.intersection(a1, a2)
        util.msgbox( a )

        a = util.difference(a1, a2)
        util.msgbox( a )

        a = util.symmetricDifference(a1, a2)
        util.msgbox( a )

    End Sub


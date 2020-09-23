Three Tier framework Readme file
================================

DATE: 03-JULY-2003 12:50
VERSION: 1.0
COPYRIGHT ©argen 2003

IF YOU WANT TO VIEW THIS WELL, PLEASE SET YOUR NOTEPAD SETTINGS TO
Lucida Console 12 Plain

Purpose:
This is an example of using  classes to retrieve information from a Data Source. Each project is a Layer in the framework:
 ------------------
|Presentation Layer| ===> uisvc.vbp (User Interface Services)
 ------------------
|Business Layer    | ===> brsvc.vbp (Business Rules Services)
 ------------------
|Data Layer        | ===> dasvc.vbp (Data Access Services)
 ------------------

The Data Interchange between layers is with XML. Is Standard and quick to pass the information with a String containing XML, instead of passing an ADO Recordset through the layers, which can be in different machines joined with a "wire".

Classes description:
--------------------
dasvc:
CAccess: This is the Class for retrieving information. Its aim is receive a Query String from the Business layer and send back a XML String with the information packaged. Also must connect and  close connection to the db, manage transactions, etc.

brsvc:
CCustomer: This is an entity class. It manage every business aspect of the Customer. It must implement the business rules for this particular entity.
Dictionary: This is a PublicNotCreatable class that contains the enums that we use to reference columns in the XML String (Offsets)
UCController: This is a controller class. Its aim is to implement every path of the UC that represent. Eg. If you have a Use Case named "Manage Customer Relationships" you build a UCManageCustRel that have all the methods necessary to implement the paths of the UC. It must translate the business inquiries into SQL queries.

uisvc:
UIController: This is a controller class that manage every aspect of the User Interface Data formatting. It manage the building of the UDTs (User Defined Types) collections and derive the inquiries of the Presentation Layer (often a Form)
Dictionary: the same as the brsvc Dictionary but with addition of the UDTs.

----------------------------------------------------------------------------------------------------

Every Method of a Class returns a Boolean value. This is because you have to tell the method caller if it was successful or not. If you need to have a return value different, you'll have to pass by reference a variable that will have that value return.

The methods calls are managed by the relationships of the classes, eg. if you have the following relationship:

                     Customer 1-----------n Order

When you ask for a Customer, this Customer have to ask for his Orders. So, from the Presentation Layer, the only thing we must do is send the message "getCustomer" to the UIController and work with the UDTs collection.


Well, this is a point if Start and possible there are many other ways to do it better, but we have to start from somewhere, haven't we?
If you know some way to do it better, please contact me or want to ask something:
argensis@yahoo.com.ar
and we will discuss it.


//show db
show dbs

// use or create a mongo db named school
use (school)

//add collection
db.createCollection(student)

//  remove database
db.dropDatabase()

//insert one document
db.student.insertOne({name:'Pole', age:32, gpa:3.0})

//insert one document
db.student.insertMany(
    {name:'Spongebob', age:28, gpa:3.2},
    {name:'Patrick', age:38, gpa:1.5},
    {name:'Sandy', age:27, gpa:4.0},
    {name:'Gary', age:18, gpa:2.5}
    )

//return all docs with a collection
db.student.find()

//data types 
db.student.insertMany(
{
    "name": "Larry",
    "age": 32,
    "gpa": 2.8,
    "fullTime": false,
    "registerDate": { "$currentDate": { "type": "date" } },
    "courses": ["Biology", "Calculus", "Chemistry"],
    "graduationDate": null,
    "address": { "street": "123 Fake St.", "city": "Bikini Bottom", "zipcode": "123" }
  }
)

//sort values
db.student.find().sort({name:-1}).limit(3)

//filter and projection
db.student.find({name:"Spongebob"})

db.student.find({gpa:4.0,fullTime:true})

db.student.find({gpa:4.0},{name:true,gpa:False})


//update field
db.student.updateOne({name:'Spongebob'},{$set:{ fullTime:true}})

//remove field
db.student.updateOne({name:'S65bca1549eeb79d1a5bdb08d'},{$unset:{ fullTime:""}})


//update many field fulltime
db.student.updateMany({},{$set:{ fullTime:False}})

//remove field fulltime
db.student.updateOne({name:'Gary'},{$unset:{ fullTime:""}})
db.student.updateOne({name:'Sandy'},{$unset:{ fullTime:""}})

//exists        //set
db.student.updateMany({fullTime:{$exists:false}},{$set:{fullTime:true}})

//DELETE
db.student.deleteMany({gpa:{$ge:4.0}})
db.student.deleteOne({name:'Larry'})

//operators                                                 #compass
//**************************************************** */
//$ne - not equal                                               
//$le - less than or equal to
//$lt - less than
//$ge - greater than or equal to                               $gte
//$gt - greater than
//$in - ["spongebob","patric"] found in
//$nin - ["spongebob","patric"] notfound in
//$and - [{fullTime:True},{age:{$lt:22}}]
//$or - [{fullTime:True},{age:{$lt:22}}]
//$nor - [{fullTime:True},{age:{$lt:22}}]
//$not - [{fullTime:True},($not:{age:{$lt:22}}}]

//***************************************************** */

//create index
db.student.createIndex({name:1})

//show index
db.student.getIndex()

//drop index
db.student.dropIndex(name_1)
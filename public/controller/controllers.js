angular.module('myApp')
.controller('StudentController', ['$scope', 'StudentService', function($scope, StudentService) {
    console.log("Hello World from controller");
var init = function() {
  $scope.newStudent={};
  $scope.studentList=[];
  StudentService.getAllStudents().then(function(students){
    angular.forEach(students,function(elem,index){
        this.push(elem);
      },$scope.studentList);
  });
};

init();

$scope.addStudent = function() {
  console.log($scope.newStudent);
  StudentService.addStudent($scope.newStudent).then(function(response){
    var newStudentRef=angular.copy($scope.newStudent);
    $scope.studentList.push(newStudentRef);
    $scope.newStudent={};
  });
};

$scope.remove = function(studentData) {
  console.log('In delete'+studentData._id);
  StudentService.deleteStudent(studentData._id).then(function(){
    $scope.studentList.splice($scope.studentList.indexOf(studentData),1);
  });
};

 

$scope.update = function(studentData) {
  console.log(studentData._id);
  StudentService.updateStudent(studentData).then(function(){
    studentData.isEditable=false;
  });

};

$scope.deselect = function() {
  $scope.student = "";
}

}])
var studentDetailsController=angular.module('myApp').controller('StudentDetailsController',['$scope','StudentService','initData',
  function($scope,StudentService,initData){
$scope.student=initData;
$scope.edit = function(student) {
  student.isEditable=true;
}; 

}]);﻿
studentDetailsController.init=['$route','$q','StudentService', function($route,$q,StudentService){
            var deferred = $q.defer();
            var studentId=$route.current.params.studentId;
            console.log('In get by student id'+studentId);
             StudentService.getStudentById(studentId).then(function(response){
              console.log('Deferred is resolved'+angular.toJson(response));
              deferred.resolve(response);
             },function(error){
              deferred.reject(error);
             });
             return deferred.promise;
          }];

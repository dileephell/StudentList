angular.module('myApp').service('StudentService',['$http','$q','routeConstants',function StudentService($http,$q,routeConstants){
var addStudentFn=function(studentdata){
	var deferred=$q.defer();
	$http.post(routeConstants.STUDENT_BASE_PREFIX,studentdata).then(function(response){
		deferred.resolve(response.data)
	},function(error){
		deferred.reject(error);
	});
	return deferred.promise;
}
var updateStudentFn=function(studentdata){
	var deferred=$q.defer();
	$http.put(routeConstants.STUDENT_BASE_PREFIX+studentdata._id,studentdata).then(function(response){
		deferred.resolve(response.data)
	},function(error){
		deferred.reject(error);
	});
	return deferred.promise;
}
var getAllStudentsFn=function(studentdata){
	var deferred=$q.defer();
	$http.get(routeConstants.STUDENT_BASE_PREFIX).then(function(response){
		deferred.resolve(response.data)
	},function(error){
		deferred.reject(error);
	});
	return deferred.promise;
}

var getStudentByIdFn=function(studentId){
	var deferred=$q.defer();
	$http.get(routeConstants.STUDENT_BASE_PREFIX+studentId).then(function(response){
		deferred.resolve(response.data)
	},function(error){
		deferred.reject(error);
	});
	return deferred.promise;
}

var deleteStudentFn=function(studentId){
	var deferred=$q.defer();
	$http.delete(routeConstants.STUDENT_BASE_PREFIX+studentId).then(function(response){
		deferred.resolve(response.data)
	},function(error){
		deferred.reject(error);
	});
	return deferred.promise;
}

this.addStudent=addStudentFn;
this.updateStudent=updateStudentFn;
this.getAllStudents=getAllStudentsFn;
this.getStudentById=getStudentByIdFn;
this.deleteStudent=deleteStudentFn;
}]);

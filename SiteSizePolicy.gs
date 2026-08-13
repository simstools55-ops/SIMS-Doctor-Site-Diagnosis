function sdsdTreatmentCapacity_(articleCount) {
  if (articleCount <= 99) return {standardMin:5, standardMax:8, hardMax:10};
  if (articleCount <= 199) return {standardMin:8, standardMax:12, hardMax:15};
  if (articleCount <= 399) return {standardMin:10, standardMax:15, hardMax:18};
  if (articleCount <= 599) return {standardMin:12, standardMax:18, hardMax:20};
  return {standardMin:15, standardMax:20, hardMax:25};
}

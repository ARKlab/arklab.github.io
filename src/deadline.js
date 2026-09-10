// Reserve time for later work; a retry must not outlive its compose event.
export function remainingTime(context={},maximumMs,reserveMs=0) {
  const remaining=context.deadlineAt===undefined?Infinity:context.deadlineAt-Date.now()-reserveMs;
  if(context.cancelled || remaining<=0) throw new Error('REQUEST_TIMEOUT');
  return Math.min(maximumMs,remaining);
}

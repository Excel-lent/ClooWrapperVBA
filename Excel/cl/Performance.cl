#if defined(cl_khr_fp64)  // Khronos extension available?
#pragma OPENCL EXTENSION cl_khr_fp64 : enable
#elif defined(cl_amd_fp64)  // AMD extension available?
#pragma OPENCL EXTENSION cl_amd_fp64 : enable
#endif

#undef MAD_4
#undef MAD_16
#undef MAD_64

#define MAD_4(x, y)     x = y*x+y;			y = x*y+x;			x = y*x+y;			y = x*y+x;
#define MAD_16(x, y)    MAD_4(x, y);        MAD_4(x, y);        MAD_4(x, y);        MAD_4(x, y);
#define MAD_64(x, y)    MAD_16(x, y);       MAD_16(x, y);       MAD_16(x, y);       MAD_16(x, y);
#define MAD_256(x,y)	MAD_64(x, y);		MAD_64(x, y);		MAD_64(x, y);		MAD_64(x, y);
#define MAD_1024(x,y)	MAD_256(x, y);		MAD_256(x, y);		MAD_256(x, y);		MAD_256(x, y);
#define MAD_4096(x,y)	MAD_1024(x, y);		MAD_1024(x, y);		MAD_1024(x, y);		MAD_1024(x, y);

__kernel void SinglePerformance(__global float* ptr, float _A) {
    float x = _A;
    float y = (float)get_local_id(0);

    MAD_1024(x, y);
    MAD_1024(x, y);

    ptr[get_global_id(0)] = y;
}

__kernel void DoublePerformance(__global double* ptr, double _A) {
	double x = _A;
	double y = (double)get_local_id(0);

	MAD_1024(x, y);
	MAD_1024(x, y);

	ptr[get_global_id(0)] = y;
}
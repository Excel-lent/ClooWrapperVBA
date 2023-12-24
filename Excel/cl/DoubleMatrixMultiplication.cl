__kernel void
DoubleMatrixMult(__global double* MResp, __global double* M1, __global double* M2, __global int* q)
{
    // Vector element index
    int i = get_global_id(0);
    int j = get_global_id(1);
    int p = get_global_size(0);
    int r = get_global_size(1);
    MResp[i + p * j] = 0;
    int QQ = q[0];
    for (int k = 0; k < QQ; k++)
    {
        MResp[i + p * j] += M1[i + p * k] * M2[k + QQ * j];
    }
}
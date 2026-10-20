
# Overview

| Skia version | Platform       | Baseline  | Case Count |
| ------------ | -------------- | --------- | ---------- |
| m144         | HO 7.0(api 23) | glesdmsaa | 769        |

Classify types of performance diff between `grdawn_vk` vs baseline(`glesdmsaa`) as  
- **Degradation**
	- **Modest**: render time diff is between [0, 1] ms
	- **Noticeable**: render time diff is between [1, 10] ms
	- **Extreme**:  render time diff is greater than 10 ms
- **Improvement**: render time diff is less than 0ms

# Distribution
---

| Category    | HO Count | Android Count |
| ----------- | -------- | ------------- |
| Degradation | 678      | 597           |
| Improvement | 89       | 160           |
| Failure     | 2        | 12            |
| In total    | 769      | 769           |

## Category - Failure
---

10x cases are skipped on Android by `SK_BUILD_FOR_ANDROID` from `DashBench` due to memory control issues.

2x cases don't have no metrics output on both HO and Android from `RepeatTileBench`

Invalid cases have render failures so their metrics outputs should be ruled out.

| Type           | cases                                                                                                                                                                                                                    | HO Count | Android Count | Description  |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------- | ------------- | ------------ |
| Skipped        | giantdashline_hori_0<br>giantdashline_hori_2<br>giantdashline_diag_0<br>giantdashline_diag_2<br>giantdashline_vert_0<br>giantdashline_vert_2<br>dashgrid_1_1_bw<br>dashgrid_1_1_aa<br>dashgrid_3_1_bw<br>dashgrid_3_1_aa | 0        | 10            |              |
| No Output      | repeatTile_BGRA_8888_X<br>repeatTile_BGRA_8888_A                                                                                                                                                                         | 2        | 2             |              |
| Render Failure | GM_paragraph__underline                                                                                                                                                                                                  | 1        | ?             | font missing |


## Category  - Degradation
---

| Sub-Category | HO Count | Android Count |
| ------------ | -------- | ------------- |
| Modest       | 448      | 451           |
| Noticeable   | 163      | 95            |
| Extreme      | 67       | 51            |

### Type - Draw Ops

|           | Modest(\[0, 1\]ms)                                | Noticable(\[1,10\]ms)                                       | Extreme(>10ms)                              |
| --------- | ------------------------------------------------- | ----------------------------------------------------------- | ------------------------------------------- |
| Rectangle | - fp cases<br>- rect shapes                       | - fp cases(blur, blendmicro, matrixconvolution)             | - long draw sequence<br>- complex fp cases  |
| Vertices  | - patches, verts, shadows(all VerticesRenderStep) | - patches(2x ops??)<br>- lines(AnalyticRRectStep)           | - long draw sequence<br>-patches(color+tex) |
| Path      | - Simple shape<br>- Path Atlas                    | - Path Atlas(hairline, bigpath)<br>- Simple shape(dashline) | - Simple shape                              |

- In extreme slower sections, 
	- Rect/FilledQuad: 
		- long draw sequence: 
			- bulk optimization: `SkCanvas::experimental_DrawEdgeAAImageSet` and `ganesh::SurfaceDrawContext::drawQuadSet`
			- non-opt, more CPU-bound: 500x (`ShapesBench`, `TiledPlaybackBench`)
		- complex fragment processor effects: blur, multiple filters
	- Vertices: 
		- long draw sequence: 
			- picture recording (from `PictureNestingBench`, 4000x )
			- `graphite::AnalyticRRectRenderStep` (`HairPointsBench`, 1000x{loop})
		- `PatchBench`: `graphite::VerticesRenderStep`(with color+tex attributes )
	- Path shape: simple shape optimization
- In slow/modest slower sections, 
	- Rect/FilledQuad: 
		- Mostly fragment process effects(blur, blending, gradient, filter, displacement ...)
		- some `AnalyticRRectRenderStep` (rotated rectangle)
	- Vertices:  
		- mostly map to `graphite::VerticesRenderStep`
		- map to `graphite::AnalyticRRectRenderStep`( from `LineBench` )
	- Path Shape:  simple shape + path atlas
	- Other analytic shape cases:
		- Oval  -> `AnalyticRRectRenderStep` (`OvalBench` & `CirclePathBench`): \[0,1\]
		- RRect -> `AnalyticRRectRenderStep` (`CommonConvexBench`):  \[1, 10\]

|          | [0,1]                                                                                                                                                                                                       | [1, 10]                                                                                                                                            | > 10 ms                                                                                                                                                             |
| -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Ganesh   | Rect(156.417)<br>FilledQuad(80.929)<br>Path(50.899) = SimpleShape(29)+InternalPath(21.899)<br>Vertices(33)<br>GlyphRunList(28)<br>Oval(17)                                                                  | FilledQuad(48.356),  <br>Path(38.5) = InternalPath(23.994)+SimpleShape(14.719) <br>Vertices(25), <br>Rect(20.984), <br>RRect(11.667)               | Vertices(16.007)<br>SimpleShape(10)<br>FilledQuad(9.578)<br>Rect(9.53)<br>Oval(6.374)                                                                               |
| Graphite | CoverBoundsRenderStep\[NonAA\](145.207)<br>AnalyticRRectRenderStep(60.948)<br>VerticesRenderStep(28 =7+11+10)<br>TessellateStrokesRenderStep(21.96)<br>BitmapTextRenderStep(15) (??)<br>PerEdgeAAQuad(8.48) | AnalyticRRectRenderStep(45.196)<br>CoverBoundsRenderStep\[NonAA\](29.177)<br>TessellateStrokesRenderStep(28.596)<br>VerticesRenderStep(16 = 4+5+7) | AnalyticRRectRenderStep(26.2)<br>PerEdgeAAQuadRenderStep(10)<br>CoverBoundsRenderStep(8)<br>StencilTessellatedWedges\[winding\](6.7)<br>VerticesRenderStep(6=1+3+2) |


- In faster section: ( **TBD** )
	- Ganesh: FilledQuad(57.568), Rect(27.432), Atlas(5)
	- Graphite: CoverBoundsRenderStep\[nonAAFill\](77.056), AnalyticRRectRenderStep(6)

### Type - Non-Draw Ops

86 nanobench cases doesn't submit draw calls. They probably measure time of recording or transfer instead.

- In extreme slower sections, 
	- Clipping
	- Formatting(Read Pixels)
- In slow/modest slower sections, 
	- Poly Utils
	- Formatting(write pixels, pre-multiple)


### Example - Top Cases
Here to list some slowest cases where diff > 10ms & ratio > 10:

| Category             | Case                                                                                                                     | Ganesh                    | Ganesh Optimization | Graphite               | Status |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------ | ------------------------- | ------------------- | ---------------------- | ------ |
| Vertices             | hair_points_mode_SrcOver_alpha_1<br>hair_points_mode_SrcOver_alpha_0.5                                                   | 1000/2000x vertices draws |                     | 4096x RRect draws      | WIP    |
| Rectangle            | bulkrect_1000_grid_uniqueimages_batch<br>bulkrect_1000_random_uniqueimages_batch<br>bulkrect_1000_grid_sharedimage_batch | 2x TextureSet draws       | batch Ops           | 1000x AA Quad draws    | DONE   |
|                      | GM_drawregion                                                                                                            | 2x Region draws           | batch Ops           | 4096x CoverBound draws | DONE   |
| Path                 | GM_dashing5_bw<br>GM_dashing5_aa                                                                                         | draw path as simple shape | Simple shape        |                        | DONE   |
| Fp - Blur            | GM_blurcircles2<br>GM_simpleblurroundrect                                                                                |                           |                     |                        | TODO   |
| Non-Draw - Formating | readpix_um_rgba_null<br>readpix_um_rgba_srgb<br>readpix_um_bgra_null<br>readpix_um_bgra_srgb                             |                           |                     |                        | TODO   |


# Summary

## Category - Degradation
---

### WIP
| Category     | case                                     | Correctness              | Perf Diff | Description | Ganesh         | Graphite                                                                          | Notes                                  |
| ------------ | ---------------------------------------- | ------------------------ | --------- | ----------- | -------------- | --------------------------------------------------------------------------------- | -------------------------------------- |
| Path         | GM_dashing5_bw                           |                          | >10ms     |             | SimpleShape    | StencilTessellatedCurvesAndTris[winding]+ StencilTessellatedWedges[winding]       | Pixel9: time diff = ; OH: diff =       |
|              | GM_convex-lineonly-paths-stroke-and-fill | quality difference       | [1,10]ms  |             | SimpleShape    | D24, add 2 ops per draw: TessellateStrokesRenderStep + TessellateWedgesRenderStep | OH: diff = 2.287 Pixel9: diff = -1.016 |
| Vertices     | patch_LOD_color_tex_xxx                  |                          | >10ms     |             | Vertices       | VerticesRenderStep                                                                |                                        |
|              |                                          |                          | [1,10]ms  |             |                |                                                                                   |                                        |
|              |                                          |                          | [0,1]ms   |             |                |                                                                                   |                                        |
| Rectangle    | game_trans_aligned_full_aa               | different result ??      | [0,1]ms   |             | FilledQuad     | CoverBoundsRenderStep\[NonAAFill\]                                                | GameBench                              |
|              |                                          |                          | [1,10]ms  |             |                |                                                                                   |                                        |
|              | GM_savelayer_with_backdrop               |                          | >10ms     |             | FilledQuad     | CoverBoundsRenderStep\[NonAAFill\]                                                |                                        |
| RRect - fp   | GM_simpleblurroundrect                   | quality difference       | >10ms     |             | 32x FilledQuad | AnalyticBlurRenderStep+AnalyticRRenderStep                                        |                                        |




### Type - DrawOps
#### Rectangle
---
- Rect: StrokeRectOp
- FilledQuad: FilledRectOp


##### Long Sequence
| case                                  | Category                              | Correctness | Ganesh                                                                                                         | Graphite                                                                                                                                                                                                                   | Notes              |
| ------------------------------------- | ------------------------------------- | ----------- | -------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------ |
| bulkrect_1000_grid_solidcolor_quad    | Long sequence, rectangle              |             | skgpu::ganesh::Device::drawEdgeAAQuad<br>`glesdmsaa`: 21.09µs<br>submission+sync:<br>`glesdmsaa`: 2.034ms      | SingleStep\[PerEdgeAAQuadRenderStep\]<br>`grdawn_vk`: 44.84µs<br>snap+submission+sync: <br> `grdawn_vk`: 22.30ms                                                                                                           | BulkRectBench      |
| tiled_playback_none_random            | Long sequence,  rectangle             |             | 5000x Device::drawRect<br>`glesdmsaa`: 24.53µs<br>submisssion+sync:<br>`glesdmsaa`: 31.52ms                    | 5000x CoverBoundsRenderStep\[NonAAFill\]<br>`grdawn_vk`: 40.16µs<br>graphite::Device::internalFlush:<br>`grdawn_vk`: 13.46ms<br>snap+(submission+sync):<br>`grdawn_vk`: 3.707ms + 162.7ms(2x CommandBuffer::addRenderPass) | TiledPlaybackBench |

##### Optimization - batching

| case                                  | Category                              | Correctness | Ganesh                                                                                                         | Graphite                                                                                                                                                    | Notes        |
| ------------------------------------- | ------------------------------------- | ----------- | -------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------ |
| GM_drawregion                         | Long sequence, batching, rectangle    |             | ganesh::Device::drawRegion                                                                                     | SkDevice::drawRegion to iterate on region list,  CoverBoundsRenderStep x {times}                                                                            | DrawRegionGM |
| bulkrect_1000_grid_uniqueimages_batch | Long sequence, batching,<br>rectangle |             | 1x SurfaceDrawContext::drawTextureSet<br>- `glesdmsaa`: 2.717ms<br>submission+sync: <br>- `glesdmsaa`: 48.44ms | 1000x SingleStep[PerEdgeAAQuadRenderStep]<br>- `grdawn_vk`: 381.8ms(1st one), 208.1µs(avg)<br>snap+(submission+sync):<br>- `grdawn_vk`:   6.433ms + 66.66ms |              |

**ImageSet/QuadSet**
in `BulkRectBench.cpp`:
- drawImagesBatch:
	`SkCanvas::experimental_DrawEdgeAAImageSet` vs `SkCanvas::drawImageRect`
- drawSolidColorBatch:
	`ganesh::SurfaceDrawContext::drawQuadSet` vs `SkCanvas::experimental_DrawEdgeAAQuad` or `SkCanvas::DrawRect`

**Region**
- `ganesh::Device::drawRegion` is specialized to create single `RegionOp`
- graphite uses the default implementation `SkDevice::drawRegion`


####  Vertices
---

##### Long Sequence

| case                               | Category                                 | Correctness | Ganesh                                                                                                                 | Graphite                                                                                                                                                                                                  | Notes                                                    |
| ---------------------------------- | ---------------------------------------- | ----------- | ---------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| picture_nesting_playback_1         | long sequence, Vertices vs AnalyticRRect |             | 9840x  DrawMeshOp<br>`SkCanvas::drawPoints`<br>- `glesdmsaa`: 18.91µs <br>submit + sync: <br>- `glesdmsaa`:  17.67ms   | 9840x AnalyticalRRectRenderStep. <br>`SkCanvas::drawPoints`<br>- `grdawn_vk`: 30.16µs<br>- `grvk`: 30µs<br>submit + sync: <br>`grdawn_vk`: 11.61ms                                                        | PictureNestingBench                                      |
| hair_points_mode_SrcOver_alpha_0.5 | long sequence, Vertices vs AnalyticRRect |             | 1000x{loop}  DrawMeshOp<br>`SkCanvas::drawPoints`<br>- `glesdmsaa`: 23.44µs<br>submit+sync:<br> - `glesdmsaa`: 21.10ms | 1000x{loop} AnalyticalRRectRenderStep<br>`SkCanvas::drawPoints`<br>- `grdawn_vk`: 8.089ms<br>- `grvk`: 7.987ms<br>snap+(submit+sync):<br>- `grdawn_vk`: 4.700ms + 22.55ms(TaskList::addCommands: 16.03ms) | HairPointsBench<br>config: (SkBlendMode::kSrcOver, 0.5f) |

#### Path Shape
---
- path atlas

##### Optimization - simple shape


#### FP effects 
---
##### Blur
- `snapRenderPass` inside `SkCanvas::drawRRect`
- `SkCanvas::attemptBlurredRRectDraw` -> `Device::drawBlurredRRect` -> `AnalyticBlurMask::MakeRRect`


### Type - NonDraw Ops
---

#### Clipping
- clip can reject draw calls.
- graphite rewrites the clip stack.

#### Format Conversion



## Category - Improvement
---

| Category   | case           | Correctness     | Perf Diff | Description | Ganesh                     | Graphite        | Notes           |
| ---------- | -------------- | --------------- | --------- | ----------- | -------------------------- | --------------- | --------------- |
| Non-DrawOp | aaclip_rect_AA | element missing | [-1, 0]ms | clipping    | 415x submissions with 0 op | no op rejected. | Good on Pixel 9 |



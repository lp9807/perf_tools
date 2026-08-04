
## Observation
---
- 100+ nanobench cases doesn't submit draw calls. They probably measure time of recording or transfer instead.
#### WIP

| case                                     | Category | Correctness              | Ganesh                     | Graphite                                                                              | Notes                                  |
| ---------------------------------------- | -------- | ------------------------ | -------------------------- | ------------------------------------------------------------------------------------- | -------------------------------------- |
| GM_convex-lineonly-paths-stroke-and-fill | path     | quality difference       | SimpleShape                | D24, add 2 ops per draw: TessellateStrokesRenderStep + TessellateWedgesRenderStep     | OH: diff = 2.287 Pixel9: diff = -1.016 |
| GM_dashing5_xx                           | path     |                          | SimpleShape                | StencilTessellatedCurvesAndTris[winding]+ StencilTessellatedWedges[winding]           | Pixel9: time diff = ; OH: diff =       |
| GM_simpleblurroundrect                   | Blur fp  | quality difference       | 32x FilledQuad             | AnalyticBlurRenderStep+AnalyticRRenderStep ( bug with counting nested snapRenderPass) |                                        |
| aaclip_rect_AA                           | clip     | element missing          | 415x submissions with 0 op | no op rejected.                                                                       | Good on Pixel 9                        |
| GM_paragraph__underline                  | font     | element missing: no font |                            |                                                                                       |                                        |
| game_trans_aligned_full_aa               | ??       | element missing          |                            |                                                                                       |                                        |
|                                          |          |                          |                            |                                                                                       |                                        |




## Distribution Analysis
---

### Overview
- In extreme slow sections, path and vertices cases take dominant places. Those are more specialized prim types for specific rendering scenarios. This means specialized optimization plays big part to the better performance of Ganesh, while no such correspondence in place yet in Graphite.
- In slow/modest slower sections, common prim types like Rect and FilledQuad starts to dominate. which makes sense but also means we need focus on other parts like shader or render pass, since they're quite simple prim types.
	- In Graphite, AnalyticRRectRenderStep seems the most popular type cross the board, except in modest slow section, where it is outnumbered by CoverBoundsRenderStep. It says as this renderer carries most of the slower paths.


### Top Cases
##### diff > 10ms & ratio > 10

| Case                    | Feature        | Optimization              | Ganesh                    | Graphite               | Status |
| ----------------------- | -------------- | ------------------------- | ------------------------- | ---------------------- | ------ |
| hair_points_mode_aa/_bw | Point          | batch Ops                 | 1000/2000x vertices draws | 4096x RRect draws      | TODO   |
| bulkrect_1000_aa/_bw    | Rect           | batch Ops                 | 2x TextureSet draws       | 1000x AA Quad draws    |        |
| GM_drawregion           | Region?        | batch Ops                 | 2x Region draws           | 4096x CoverBound draws | TODO   |
| readpix_bgra/rgba_xx    | Format/Trasfer |                           |                           |                        | TODO   |
| GM_blurcircles2         | Blur           |                           |                           |                        | WIP    |
| GM_dashing_xx           | Path           | draw path as simple shape |                           |                        | WIP    |

### Extreme Slower
##### diff > 10ms

|          | Summary                                                                                                                                                             |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Ganesh   | Vertices(16.007)<br>SimpleShape(10)<br>FilledQuad(9.578)<br>Rect(9.53)<br>Oval(6.374)                                                                               |
| Graphite | AnalyticRRectRenderStep(26.2)<br>PerEdgeAAQuadRenderStep(10)<br>CoverBoundsRenderStep(8)<br>StencilTessellatedWedges\[winding\](6.7)<br>VerticesRenderStep(6=1+3+2) |

- In Ganesh, top used primitive type is vertices, which is related batch optimization, i.e. batch large list of other type of primitives into less times of bigger vertice draw calls. The corresponding primitive type in Graphite is VerticesRenderStep type, which is used but less frequent at 5th position. The performance of single vertice draw call in Graphite Vulkan also is worth of digging.
- 2nd place is SimpleShape in Ganesh, that is an optimized path of path renderer so to bypass atlas rendering. In this section most path-related cases are done in this path.
- Then comes to FilledQuad and Rect, which are most common types. Perhaps more related to shader,  batch optimization or other fp configs. 


#### ratio > 10

### Slower
##### diff in [1, 10] ms

| Backend  | Summary                                                                                                                                            |
| :------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| Ganesh   | FilledQuad(48.356),  <br>Path(38.5) = InternalPath(23.994)+SimpleShape(14.719) <br>Vertices(25), <br>Rect(20.984), <br>RRect(11.667)               |
| Graphite | AnalyticRRectRenderStep(45.196)<br>CoverBoundsRenderStep\[NonAA\](29.177)<br>TessellateStrokesRenderStep(28.596)<br>VerticesRenderStep(16 = 4+5+7) |


##### ratio in [2, 10]

### Modest Slower
##### diff  in [0, 1]

| Backend  | Summary                                                                                                                                                                                                     |
| -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Ganesh   | Rect(156.417)<br>FilledQuad(80.929)<br>Path(50.899) = SimpleShape(29)+InternalPath(21.899)<br>Vertices(33)<br>GlyphRunList(28)<br>Oval(17)                                                                  |
| Graphite | CoverBoundsRenderStep\[NonAA\](145.207)<br>AnalyticRRectRenderStep(60.948)<br>VerticesRenderStep(28 =7+11+10)<br>TessellateStrokesRenderStep(21.96)<br>BitmapTextRenderStep(15) (??)<br>PerEdgeAAQuad(8.48) |
Most slow cases take in modest slower section, so this is rather representative in term of performance comparison. 
Ganesh:
- Top Ganesh primitive types are Rect and FilledQuad, which just says that those 2 are most often used types. However, need more information, e.g. shader config, draw times to conduct more meaningful comparison.
- Path and Vertices takes next 2 positions, which shows their performance on Graphite is generally slower.
- 5th place is taken by text.
Graphite side, 
- The order oughly can match Ganesh. 
- 2nd place is AnalyticRRectRenderStep, which seems correspond to FilledQuad? Need further investigation.


##### ratio in [1, 2]

### Faster
Ganesh: FilledQuad(57.568), Rect(27.432), Atlas(5)
Graphite: CoverBoundsRenderStep\[nonAAFill\](77.056), AnalyticRRectRenderStep(6)

## Summary
---
### Feature : Path
- path atlas
- Specialized Optimization: simple shape


### Clip
- clip can reject draw calls.
- graphite rewrites the clip stack.

### Feature:  Batch Op Optimization

#### Vertices
??

#### Rect-related
??


### Feature: Blur
- snapRenderPass inside SkCanvas::drawRRect
- SkCanvas::attemptBlurredRRectDraw -> Device::drawBlurredRRect -> AnalyticBlurMask::MakeRRect


### Feature: Format Conversion/Transfer

TBA

function extract_fig_mesh_color(figPath, outMat)
% EXTRACT_FIG_MESH_COLOR  Pull geometry + displayed colors out of a MATLAB .fig.
%
%   extract_fig_mesh_color(FIGPATH, OUTMAT) opens the saved figure FIGPATH
%   (a colored surface mesh, e.g. Cell_1_min_curv.fig), reads the mesh patch,
%   reproduces the RGB colors the figure displays (mapping the scalar
%   FaceVertexCData through the axes Colormap and CLim, exactly as MATLAB
%   renders scaled CData), and saves V, F, RGB and a PERFACE flag to the -v7
%   MAT-file OUTMAT for the Python glTF baker.
%
%   V       Nx3  vertex positions
%   F       Mx3  1-based triangle vertex indices
%   RGB     Kx3  colors in [0,1]; K==M when PERFACE==1, else K==N
%   PERFACE 1 if colors are per-face (FaceColor 'flat'), 0 if per-vertex.
    f = openfig(figPath, 'invisible');
    cleaner = onCleanup(@() close(f));

    p = findobj(f, 'Type', 'patch');
    if isempty(p)
        error('extract_fig_mesh_color:noPatch', 'no patch in %s', figPath);
    end
    % The nucleus mesh is the patch with the most faces (ignore marker patches).
    [~, k] = max(arrayfun(@(h) size(h.Faces, 1), p));
    p = p(k);

    V = double(p.Vertices);
    F = double(p.Faces);
    c = p.FaceVertexCData;
    ax = ancestor(p, 'axes');
    cm = colormap(ax);
    n = size(cm, 1);

    if size(c, 2) == 3
        % Already truecolor RGB (per-face or per-vertex).
        RGB = double(c);
    else
        % Indexed scalar -> map through the colormap using CLim, matching how
        % MATLAB renders scaled CData (floor into n bins, clamped).
        c = double(c(:));
        cl = double(ax.CLim);
        t = (c - cl(1)) / (cl(2) - cl(1));
        t = min(max(t, 0), 1);
        ix = floor(t * n) + 1;
        ix = min(max(ix, 1), n);
        RGB = cm(ix, :);
    end

    perface = double(size(RGB, 1) == size(F, 1));  %#ok<NASGU>
    save(outMat, 'V', 'F', 'RGB', 'perface', '-v7');
end
